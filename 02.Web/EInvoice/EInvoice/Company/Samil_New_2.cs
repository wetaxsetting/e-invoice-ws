using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Xml;

using Oracle.DataAccess.Client;
using Oracle.DataAccess.Types;
//using Oracle.ManagedDataAccess.Client;

namespace EInvoice.Company
{
    public class Samil_New_3
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
            string read_prive = "", read_en = "", read_amount = "", amount_vat = "", amount_total ="", amount_trans = "", amount_net="", lb_amount_trans = "";
            

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


            /*if (dt.Rows[0]["vatamount_display"].ToString().Trim() == "0")
            {
                amount_vat = "-";
            }
            else
            {
                amount_vat = dt.Rows[0]["vatamount_display"].ToString();
            }*/

            //read_en = dt.Rows[0]["TotalAmountInWord"].ToString();
            int end = 0;
            int count = count_page_v + r;
            double height = 130;
            StringBuilder htmlStr = new StringBuilder("");
            string heigh = "", heigh_d = "";


            htmlStr.Append("<!DOCTYPE html PUBLIC -//W3C//DTD HTML 4.01 Transitional//EN http://www.w3.org/TR/html4/loose.dtd>																\n");
            htmlStr.Append("<html>                                                                                                                                                          \n");
            htmlStr.Append("<head>                                                                                                                                                          \n");
            htmlStr.Append("<meta http-equiv=Content-Type content=text/html; charset=UTF-8>                                                                                                 \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append("<script type=text/javascript                                                                                                                                    \n");
            htmlStr.Append("	src=${pageContext.request.contextPath}/system/syscommand.js></script>                                                                                       \n");
            htmlStr.Append("<title>Report E-Invoice</title>                                                                                                                                 \n");
            htmlStr.Append("<!-- Set page size here: A5, A4 or A3 -->                                                                                                                       \n");
            htmlStr.Append("<!-- Set also landscape if you need -->                                                                                                                         \n");
            htmlStr.Append("<style>                                                                                                                                                         \n");
            htmlStr.Append("@page {                                                                                                                                                         \n");
            htmlStr.Append("	size: A4                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("</style>                                                                                                                                                        \n");
            htmlStr.Append("<link href='https://fonts.googleapis.com/css?family=Tangerine:700'                                                                                              \n");
            htmlStr.Append("	rel='stylesheet' type='text/css'>                                                                                                                           \n");
            htmlStr.Append("<style>                                                                                                                                                         \n");
            htmlStr.Append("/*body   { font-family: serif }                                                                                                                                 \n");
            htmlStr.Append("    h1     { font-family: 'Tangerine', cursive; font-size: 40pt; line-height: 18mm}                                                                             \n");
            htmlStr.Append("    h2, h3 { font-family: 'Tangerine', cursive; font-size: 24pt; line-height: 7mm }                                                                             \n");
            htmlStr.Append("    h4     { font-size: 16.25pt; line-height: 1mm }                                                                                                                \n");
            htmlStr.Append("    h2 + p { font-size: 22.5pt; line-height: 7mm }                                                                                                                \n");
            htmlStr.Append("    h3 + p { font-size: 17.5pt; line-height: 7mm }                                                                                                                \n");
            htmlStr.Append("    li     { font-size: 13.75pt; line-height: 5mm }                                                                                                                \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append("    h1      { margin: 0 }                                                                                                                                       \n");
            htmlStr.Append("    h1 + ul { margin: 2mm 0 5mm }                                                                                                                               \n");
            htmlStr.Append("    h2, h3  { margin: 0 3mm 3mm 0; float: left }                                                                                                                \n");
            htmlStr.Append("    h2 + p,                                                                                                                                                     \n");
            htmlStr.Append("    h3 + p  { margin: 0 0 3mm 50mm }                                                                                                                            \n");
            htmlStr.Append("    //h4      { margin: 1mm 0 0 2mm; border-bottom: 1px solid black }                                                                                           \n");
            htmlStr.Append("    h4 + ul { margin: 5mm 0 0 50mm }                                                                                                                            \n");
            htmlStr.Append("    article { border: 4px double black; padding: 5mm 10mm; border-radius: 3mm }*/                                                                               \n");
            htmlStr.Append("body {                                                                                                                                                          \n");
            htmlStr.Append("	color: blue;                                                                                                                                                \n");
            htmlStr.Append("	font-size: 100%;                                                                                                                                            \n");
            htmlStr.Append("	background-image: url(assets/Solution.jpg);                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append("h1 {                                                                                                                                                            \n");
            htmlStr.Append("	color: #00FF00;                                                                                                                                             \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append("p {                                                                                                                                                             \n");
            htmlStr.Append("	color: rgb(0, 0, 255)                                                                                                                                       \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append("headline1 {                                                                                                                                                     \n");
            htmlStr.Append("	background-image: url(assets/Solution.jpg);                                                                                                                 \n");
            htmlStr.Append("	background-repeat: no-repeat;                                                                                                                               \n");
            htmlStr.Append("	background-position: left top;                                                                                                                              \n");
            htmlStr.Append("	padding-top: 68px;                                                                                                                                          \n");
            htmlStr.Append("	margin-bottom: 50px;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append("headline2 {                                                                                                                                                     \n");
            htmlStr.Append("	background-image: url(images/newsletter_headline2.gif);                                                                                                     \n");
            htmlStr.Append("	background-repeat: no-repeat;                                                                                                                               \n");
            htmlStr.Append("	background-position: left top;                                                                                                                              \n");
            htmlStr.Append("	padding-top: 68px;                                                                                                                                          \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append("<!--                                                                                                                                                            \n");
            htmlStr.Append("table {                                                                                                                                                         \n");
            htmlStr.Append("	mso-displayed-decimal-separator: \\.;                                                                                                                        \n");
            htmlStr.Append("	mso-displayed-thousand-separator: \\,;                                                                                                                       \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".font527974 {                                                                                                                                                   \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 11.25pt;                                                                                                                                           \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Times New Roman, serif;                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".font627974 {                                                                                                                                                   \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Times New Roman, serif;                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".font727974 {                                                                                                                                                   \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 11.25pt;                                                                                                                                           \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".font827974 {                                                                                                                                                   \n");
            htmlStr.Append("	color: black;                                                                                                                                               \n");
            htmlStr.Append("	font-size: 11.25pt;                                                                                                                                           \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".font927974 {                                                                                                                                                   \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".font1027974 {                                                                                                                                                  \n");
            htmlStr.Append("	color: black;                                                                                                                                               \n");
            htmlStr.Append("	font-size: 11.25pt;                                                                                                                                           \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".font1127974 {                                                                                                                                                  \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 11.875pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".font1227974 {                                                                                                                                                  \n");
            htmlStr.Append("	color: red;                                                                                                                                                 \n");
            htmlStr.Append("	font-size: 17.0pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".font1327974 {                                                                                                                                                  \n");
            htmlStr.Append("	color: red;                                                                                                                                                 \n");
            htmlStr.Append("	font-size: 18.0pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Times New Roman, serif;                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".font1427974 {                                                                                                                                                  \n");
            htmlStr.Append("	color: red;                                                                                                                                                 \n");
            htmlStr.Append("	font-size: 18.0pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".font1527974 {                                                                                                                                                  \n");
            htmlStr.Append("	color: red;                                                                                                                                                 \n");
            htmlStr.Append("	font-size: 18.0pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".font1627974 {                                                                                                                                                  \n");
            htmlStr.Append("	color: red;                                                                                                                                                 \n");
            htmlStr.Append("	font-size: 11.25pt;                                                                                                                                           \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Times New Roman, serif;                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".font1727974 {                                                                                                                                                  \n");
            htmlStr.Append("	color: red;                                                                                                                                                 \n");
            htmlStr.Append("	font-size: 11.25pt;                                                                                                                                           \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Times New Roman, serif;                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".font1827974 {                                                                                                                                                  \n");
            htmlStr.Append("	color: red;                                                                                                                                                 \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                           \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Times New Roman, serif;                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".font1927974 {                                                                                                                                                  \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".font2027974 {                                                                                                                                                  \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Century, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".font2127974 {                                                                                                                                                  \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Times New Roman, serif;                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".font2227974 {                                                                                                                                                  \n");
            htmlStr.Append("	color: #0066CC;                                                                                                                                             \n");
            htmlStr.Append("	font-size: 11.875pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Times New Roman, serif;                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl6527974 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 11.25pt;                                                                                                                                           \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: general;                                                                                                                                        \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl6627974 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 11.25pt;                                                                                                                                           \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: general;                                                                                                                                        \n");
            htmlStr.Append("	vertical-align: mid;                                                                                                                                        \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl6727974 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 11.875pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: 2.0pt double windowtext;                                                                                                                       \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl6827974 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 11.875pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl6927974 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: 2.0pt double windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl7027974 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: navy;                                                                                                                                                \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: general;                                                                                                                                        \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl7127974 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: navy;                                                                                                                                                \n");
            htmlStr.Append("	font-size: 24.0pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: general;                                                                                                                                        \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl7227974 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: black;                                                                                                                                               \n");
            htmlStr.Append("	font-size: 11.25pt;                                                                                                                                           \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: general;                                                                                                                                        \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl7327974 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: black;                                                                                                                                               \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: general;                                                                                                                                        \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 2.0pt double windowtext;                                                                                                                        \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: 2.0pt double windowtext;                                                                                                                       \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl7427974 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: black;                                                                                                                                               \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: general;                                                                                                                                        \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 2.0pt double windowtext;                                                                                                                        \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl7527974 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: navy;                                                                                                                                                \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: left;                                                                                                                                           \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 2.0pt double windowtext;                                                                                                                        \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl7627974 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: navy;                                                                                                                                                \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: left;                                                                                                                                           \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 2.0pt double windowtext;                                                                                                                        \n");
            htmlStr.Append("	border-right: 2.0pt double windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl7727974 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: black;                                                                                                                                               \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: general;                                                                                                                                        \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: 2.0pt double windowtext;                                                                                                                       \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl7827974 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: navy;                                                                                                                                                \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: left;                                                                                                                                           \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl7927974 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: navy;                                                                                                                                                \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: left;                                                                                                                                           \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: 2.0pt double windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl8027974 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 11.25pt;                                                                                                                                           \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: left;                                                                                                                                           \n");
            htmlStr.Append("	vertical-align: mid;                                                                                                                                        \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: 2.0pt double windowtext;                                                                                                                       \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl8127974 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 1pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: general;                                                                                                                                        \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: 2.0pt double windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl8227974 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 11.25pt;                                                                                                                                           \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: general;                                                                                                                                        \n");
            htmlStr.Append("	vertical-align: mid;                                                                                                                                        \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: 2.0pt double windowtext;                                                                                                                       \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl82279741 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 4.0pt;                                                                                                                                           \n");
            htmlStr.Append("	font-weight: 100;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: general;                                                                                                                                        \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: 2.0pt double windowtext;                                                                                                                       \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl8327974 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: general;                                                                                                                                        \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                       \n");
            htmlStr.Append("	border-left: 2.0pt double windowtext;                                                                                                                       \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl8427974 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: general;                                                                                                                                        \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                       \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl8527974 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: general;                                                                                                                                        \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: 2.0pt double windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                       \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl8627974 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 11.25pt;                                                                                                                                           \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                          \n");
            htmlStr.Append("	border-right: 1pt solid windowtext;                                                                                                                        \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                       \n");
            htmlStr.Append("	border-left: 2.0pt double windowtext;                                                                                                                       \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl8727974 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 11.25pt;                                                                                                                                           \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                          \n");
            htmlStr.Append("	border-right: 2.0pt double windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                       \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                         \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl8827974 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: general;                                                                                                                                        \n");
            htmlStr.Append("	vertical-align: midle;                                                                                                                                      \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                          \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                       \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl8927974 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 12.5pt;                                                                                                                                           \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: right;                                                                                                                                          \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                          \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                       \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl9027974 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: general;                                                                                                                                        \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                          \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl9127974 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 12.50pt;                                                                                                                                           \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: general;                                                                                                                                        \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                          \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: 2.0pt double windowtext;                                                                                                                       \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl9227974 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 11.875pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: left;                                                                                                                                           \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: 1pt dotted windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl9327974 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 11.8pt;                                                                                                                                           \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: left;                                                                                                                                           \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: 2.0pt double windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-bottom: 1pt dotted windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append(".xl93279743 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 11.6pt;                                                                                                                                           \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: left;                                                                                                                                           \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: 2.0pt double windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-bottom: 1pt dotted windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");

            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl9427974 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: general;                                                                                                                                        \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                          \n");
            htmlStr.Append("	border-right: 2.0pt double windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl9527974 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: right;                                                                                                                                          \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                          \n");
            htmlStr.Append("	border-right: 2.0pt double windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                       \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                         \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl9627974 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 11.25pt;                                                                                                                                           \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: right;                                                                                                                                          \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: 2.0pt double windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-bottom: 1pt hairline windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                         \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl9727974 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 12.5pt;                                                                                                                                            \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: right;                                                                                                                                          \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt hairline windowtext;                                                                                                                       \n");
            htmlStr.Append("	border-right: 2.0pt double windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-bottom: 1pt hairline windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                         \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append(".xl97279741 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 12.5pt;                                                                                                                                            \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: right;                                                                                                                                          \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt hairline windowtext;                                                                                                                       \n");
            htmlStr.Append("	border-right: 2.0pt double windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                         \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");

            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl9827974 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 11.25pt;                                                                                                                                           \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt hairline windowtext;                                                                                                                       \n");
            htmlStr.Append("	border-right: 1pt solid windowtext;                                                                                                                        \n");
            htmlStr.Append("	border-bottom: 1pt hairline windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-left: 2.0pt double windowtext;                                                                                                                       \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append(".xl98279741 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 11.25pt;                                                                                                                                           \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt hairline windowtext;                                                                                                                       \n");
            htmlStr.Append("	border-right: 1pt solid windowtext;                                                                                                                        \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-left: 2.0pt double windowtext;                                                                                                                       \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl9927974 {                                                                                                                                                    \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 11.875pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: left;                                                                                                                                           \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: 2.0pt double windowtext;                                                                                                                       \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl10027974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: left;                                                                                                                                           \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl10127974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl10227974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 11.875pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                          \n");
            htmlStr.Append("	border-right: 1pt solid windowtext;                                                                                                                        \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: 2.0pt double windowtext;                                                                                                                       \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl10327974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 11.875pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                          \n");
            htmlStr.Append("	border-right: 2.0pt double windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                         \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl10427974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: 1pt solid windowtext;                                                                                                                        \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                       \n");
            htmlStr.Append("	border-left: 2.0pt double windowtext;                                                                                                                       \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl10527974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 11.875pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: 2.0pt double windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                       \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                         \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl10627974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: #0070C0;                                                                                                                                             \n");
            htmlStr.Append("	font-size: 12.5pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Times New Roman, serif;                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: general;                                                                                                                                        \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl10727974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 12.5pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Times New Roman, serif;                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: general;                                                                                                                                        \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl10827974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: red;                                                                                                                                                 \n");
            htmlStr.Append("	font-size: 11.25pt;                                                                                                                                           \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Times New Roman, serif;                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: general;                                                                                                                                        \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                          \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                         \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl10927974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Times New Roman, serif;                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: general;                                                                                                                                        \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                          \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl11027974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Times New Roman, serif;                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: general;                                                                                                                                        \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                          \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl11127974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Times New Roman, serif;                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: general;                                                                                                                                        \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                          \n");
            htmlStr.Append("	border-right: 2.0pt double windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl11227974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: red;                                                                                                                                                 \n");
            htmlStr.Append("	font-size: 11.25pt;                                                                                                                                           \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Times New Roman, serif;                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: general;                                                                                                                                        \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                       \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                         \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl11327974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: red;                                                                                                                                                 \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Times New Roman, serif;                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: general;                                                                                                                                        \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                       \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl11427974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: red;                                                                                                                                                 \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Times New Roman, serif;                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: general;                                                                                                                                        \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: 2.0pt double windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                       \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl11527974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: red;                                                                                                                                                 \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: left;                                                                                                                                           \n");
            htmlStr.Append("	vertical-align: mid;                                                                                                                                        \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: 2.0pt double windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl11627974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: #0070C0;                                                                                                                                             \n");
            htmlStr.Append("	font-size: 11.875pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Times New Roman, serif;                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: general;                                                                                                                                        \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: 2.0pt double windowtext;                                                                                                                       \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl11727974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 11.875pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Times New Roman, serif;                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: general;                                                                                                                                        \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: 2.0pt double windowtext;                                                                                                                       \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl11827974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: red;                                                                                                                                                 \n");
            htmlStr.Append("	font-size: 11.25pt;                                                                                                                                           \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Times New Roman, serif;                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: left;                                                                                                                                           \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                         \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                        \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                              \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl11927974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: red;                                                                                                                                                 \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Times New Roman, serif;                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: left;                                                                                                                                           \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                              \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl12027974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: red;                                                                                                                                                 \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Times New Roman, serif;                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: left;                                                                                                                                           \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: 2.0pt double windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                              \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl12127974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 11.875pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: 2.0pt double windowtext;                                                                                                                     \n");
            htmlStr.Append("	border-left: 2.0pt double windowtext;                                                                                                                       \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl12227974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: 2.0pt double windowtext;                                                                                                                     \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl12327974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: 2.0pt double windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-bottom: 2.0pt double windowtext;                                                                                                                     \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl12427974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 11.875pt;                                                                                                                                           \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                          \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl12527974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 11.25pt;                                                                                                                                           \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: left;                                                                                                                                           \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: 1pt dotted windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-left: 2.0pt double windowtext;                                                                                                                       \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl12627974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 11.875pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                          \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: 2.0pt double windowtext;                                                                                                                       \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl12727974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                          \n");
            htmlStr.Append("	border-right: 2.0pt double windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl12827974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 11.875pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: 2.0pt double windowtext;                                                                                                                       \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl12927974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 11.875pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl13027974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 11.875pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl13127974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: 2.0pt double windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl13227974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 12pt;                                                                                                                                             \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: left;                                                                                                                                           \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                          \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                       \n");
            htmlStr.Append("	border-left: 2.0pt double windowtext;                                                                                                                       \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl13327974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: left;                                                                                                                                           \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                          \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                       \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl13427974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: left;                                                                                                                                           \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                          \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                       \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                          \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl13527974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: right;                                                                                                                                          \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                          \n");
            htmlStr.Append("	border-right: 1pt solid windowtext;                                                                                                                        \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                       \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl13627974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                           \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: left;                                                                                                                                           \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt hairline windowtext;                                                                                                                       \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                         \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append(".xl136279741 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                           \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: left;                                                                                                                                           \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt hairline windowtext;                                                                                                                       \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                         \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");

            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl13727974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: left;                                                                                                                                           \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt hairline windowtext;                                                                                                                       \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("    border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl13827974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: left;                                                                                                                                           \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt hairline windowtext;                                                                                                                       \n");
            htmlStr.Append("	border-right: 1pt solid windowtext;                                                                                                                        \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl13927974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 12.5pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt hairline windowtext;                                                                                                                       \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                         \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append(".xl139279741 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 12.5pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt hairline windowtext;                                                                                                                       \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                         \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");

            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl14027974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt hairline windowtext;                                                                                                                       \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl14127974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 12.5pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: right;                                                                                                                                          \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt hairline windowtext;                                                                                                                       \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                         \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append(".xl141279741 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 12.5pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: right;                                                                                                                                          \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt hairline windowtext;                                                                                                                       \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                         \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl14227974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: right;                                                                                                                                          \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt hairline windowtext;                                                                                                                       \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl14327974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 11.875pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: 1pt solid windowtext;                                                                                                                        \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                       \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                         \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl14427974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: 1pt solid windowtext;                                                                                                                        \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                       \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                         \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl14527974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 11.25pt;                                                                                                                                           \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border: 1pt solid windowtext;                                                                                                                              \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl14627974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 11.25pt;                                                                                                                                           \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border: 1pt solid windowtext;                                                                                                                              \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl14727974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 12.5pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: justify;                                                                                                                                    \n");
            htmlStr.Append("	border-top: 1pt dotted windowtext;                                                                                                                         \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl14827974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: justify;                                                                                                                                    \n");
            htmlStr.Append("	border-top: 1pt dotted windowtext;                                                                                                                         \n");
            htmlStr.Append("	border-right: 2.0pt double windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl14927974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: justify;                                                                                                                                    \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: 1pt dotted windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl15027974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: justify;                                                                                                                                    \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: 2.0pt double windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-bottom: 1pt dotted windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl15127974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 11.875pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                          \n");
            htmlStr.Append("	border-right: 1pt solid windowtext;                                                                                                                        \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                         \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl15227974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: left;                                                                                                                                           \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: 1pt dotted windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                              \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl15327974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: left;                                                                                                                                           \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: 2.0pt double windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-bottom: 1pt dotted windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                              \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl15427974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: 1pt dotted windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl15527974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: left;                                                                                                                                           \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt dotted windowtext;                                                                                                                         \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: 1pt dotted windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                              \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl15627974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: left;                                                                                                                                           \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt dotted windowtext;                                                                                                                         \n");
            htmlStr.Append("	border-right: 2.0pt double windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-bottom: 1pt dotted windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                              \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl15727974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: black;                                                                                                                                               \n");
            htmlStr.Append("	font-size: 12.5pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                          \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: 2.0pt double windowtext;                                                                                                                       \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl15827974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: black;                                                                                                                                               \n");
            htmlStr.Append("	font-size: 24.0pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                          \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl15927974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: black;                                                                                                                                               \n");
            htmlStr.Append("	font-size: 24.0pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                     \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                          \n");
            htmlStr.Append("	border-right: 2.0pt double windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                        \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl16027974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                       \n");
            htmlStr.Append("	border-left: 2.0pt double windowtext;                                                                                                                       \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl16127974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                       \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl16227974 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                          \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: center;                                                                                                                                         \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: 2.0pt double windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                       \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append(".xl93279742 {                                                                                                                                                   \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                        \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                          \n");
            htmlStr.Append("	font-size: 11.875pt;                                                                                                                                           \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                           \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                         \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                      \n");
            htmlStr.Append("	font-family: Cambria, serif;                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                        \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                 \n");
            htmlStr.Append("	text-align: left;                                                                                                                                           \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                     \n");
            htmlStr.Append("	border-top: none;                                                                                                                                           \n");
            htmlStr.Append("	border-right: none;                                                                                                                                         \n");
            htmlStr.Append("	border-bottom: 1pt dotted windowtext;                                                                                                                      \n");
            htmlStr.Append("	border-left: none;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                          \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                    \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                               \n");
            htmlStr.Append(".xl99279741                                                                                                                                                     \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                               \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                         \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                           \n");
            htmlStr.Append("	font-size: 11.875pt;                                                                                                                                           \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                            \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                          \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                       \n");
            htmlStr.Append("	font-family:Cambria, serif;                                                                                                                                 \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                         \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                  \n");
            htmlStr.Append("	text-align:left;                                                                                                                                            \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                      \n");
            htmlStr.Append("	border-right:none;                                                                                                                                          \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                         \n");
            htmlStr.Append("	border-left:none;                                                                                                                                           \n");
            htmlStr.Append("	background:white;                                                                                                                                           \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                     \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                        \n");
            htmlStr.Append("-->                                                                                                                                                             \n");
            htmlStr.Append("</style>                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append("</head>                                                                                                                                                                     \n");
            htmlStr.Append("<body class='A4'>	                                                                                                                                                        \n");

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

            double v_totalHeightLastPage = 258.5;// 300;// 258.5;

            double v_totalHeightPage = 580;//   540;

            for (int k = 0; k < v_countNumberOfPages; k++)
            {
                v_totalHeightPage = 500;// 540;

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
                htmlStr.Append("		<table border=0 cellpadding=0 cellspacing=0 width=742 class=xl6627974                                                                                   \n");
                htmlStr.Append("			style='border-collapse: collapse; table-layout: fixed; width: 530pt'>                                                                               \n");
                htmlStr.Append("			<col class=xl6627974 width=59                                                                                                                       \n");
                htmlStr.Append("				style='mso-width-source: userset; mso-width-alt: 2104; width: 40.00pt'>                                                                            \n");
                htmlStr.Append("			<!-- 15 -->                                                                                                                                         \n");
                htmlStr.Append("			<col class=xl6627974 width=135                                                                                                                      \n");
                htmlStr.Append("				style='mso-width-source: userset; mso-width-alt: 4807; width: 70.00pt'>                                                                            \n");
                htmlStr.Append("			<!-- 14 -->                                                                                                                                         \n");
                htmlStr.Append("			<col class=xl6627974 width=115                                                                                                                      \n");
                htmlStr.Append("				style='mso-width-source: userset; mso-width-alt: 4096; width: 38.75pt'>           <!--28.75 -->                                                                 \n");
                htmlStr.Append("			<!-- 13 -->                                                                                                                                         \n");
                htmlStr.Append("			<col class=xl6627974 width=40                                                                                                                       \n");
                htmlStr.Append("				style='mso-width-source: userset; mso-width-alt: 1422; width: 66.25pt'>           <!--46.25 -->                                                                 \n");
                htmlStr.Append("			<!-- 12 -->                                                                                                                                         \n");
                htmlStr.Append("			<col class=xl6627974 width=29                                                                                                                       \n");
                htmlStr.Append("				style='mso-width-source: userset; mso-width-alt: 1024; width: 47.5pt'>              <!--52.5 -->                                                              \n");
                htmlStr.Append("			<!-- 11 -->                                                                                                                                         \n");
                htmlStr.Append("			<col class=xl6627974 width=34                                                                                                                       \n");
                htmlStr.Append("				style='mso-width-source: userset; mso-width-alt: 1223; width: 32.5pt'>             <!-- 37.5 -->                                                               \n");
                htmlStr.Append("			<!-- 10 -->                                                                                                                                         \n");
                htmlStr.Append("			<col class=xl6627974 width=25                                                                                                                       \n");
                htmlStr.Append("				style='mso-width-source: userset; mso-width-alt: 881; width: 43.75pt'>                                                                             \n");
                htmlStr.Append("			<!-- 9 -->                                                                                                                                          \n");
                htmlStr.Append("			<col class=xl6627974 width=202                                                                                                                      \n");
                htmlStr.Append("				style='mso-width-source: userset; mso-width-alt: 7168; width: 53.75pt'>           <!--63.75 -->                                                                 \n");
                htmlStr.Append("			<!-- 8 -->                                                                                                                                          \n");
                htmlStr.Append("			<col class=xl6627974 width=39                                                                                                                       \n");
                htmlStr.Append("				style='mso-width-source: userset; mso-width-alt: 1393; width: 26.25pt'>                                                                            \n");
                htmlStr.Append("			<!-- 7 -->                                                                                                                                          \n");
                htmlStr.Append("			<col class=xl6627974 width=30                                                                                                                       \n");
                htmlStr.Append("				style='mso-width-source: userset; mso-width-alt: 1080; width: 22.5pt'>                                                                            \n");
                htmlStr.Append("			<!-- 6 -->                                                                                                                                          \n");
                htmlStr.Append("			<col class=xl6627974 width=63                                                                                                                       \n");
                htmlStr.Append("				style='mso-width-source: userset; mso-width-alt: 2247; width: 37.5pt'>                                                                            \n");
                htmlStr.Append("			<!-- 5 -->                                                                                                                                          \n");
                htmlStr.Append("			<col class=xl6627974 width=34                                                                                                                       \n");
                htmlStr.Append("				style='mso-width-source: userset; mso-width-alt: 1223; width: 32.5pt'>                                                                            \n");
                htmlStr.Append("			<!-- 4 -->                                                                                                                                          \n");
                htmlStr.Append("			<col class=xl6627974 width=49                                                                                                                       \n");
                htmlStr.Append("				style='mso-width-source: userset; mso-width-alt: 1735; width:27.5pt'>                                                                            \n");
                htmlStr.Append("			<!-- 3 -->                                                                                                                                          \n");
                htmlStr.Append("			<col class=xl6627974 width=38                                                                                                                       \n");
                htmlStr.Append("				style='mso-width-source: userset; mso-width-alt: 1336; width:42.5pt'>                                                                            \n");
                htmlStr.Append("			<!-- 2 -->                                                                                                                                          \n");
                htmlStr.Append("			<col class=xl6627974 width=144                                                                                                                      \n");
                htmlStr.Append("				style='mso-width-source: userset; mso-width-alt: 5120; width: 73.75pt'>                                                                            \n");
                htmlStr.Append("			<!-- 1 -->                                                                                                                                          \n");
                htmlStr.Append("			<tr class=xl7227974 height=40                                                                                                                       \n");
                htmlStr.Append("				style='mso-height-source: userset; height: 22.5pt'>                                                                                              \n");
                htmlStr.Append("				<td height=40 colspan=2 class=xl7727974                                                                                                         \n");
                htmlStr.Append("					style='height: 22.5pt; border-top: 2.0pt double windowtext;'>&nbsp;</td>                                                                      \n");
                htmlStr.Append("				<td colspan=7 rowspan=1 height=50 class=xl7227974 width=445                                                                                     \n");
                htmlStr.Append("		style='mso-ignore: colspan-rowspan; height: 25pt; width: 275pt; border-top: 2.0pt double windowtext; vertical-align: middle;'><![if !vml]><span         \n");
                htmlStr.Append("					style='mso-ignore: vglayout'>                                                                                                               \n");
                htmlStr.Append("						<table cellpadding=0 cellspacing=0>                                                                                                     \n");
                htmlStr.Append("							<tr>                                                                                                                                \n");
                htmlStr.Append("								<td width=28></td>                                                                                                              \n");
                htmlStr.Append("								<td style='vertical-align: mid;'><img width=362 height=43.25                                                                      \n");
                htmlStr.Append("									src=D:\\webproject\\e-invoice-ws\\02.Web\\EInvoice\\img\\Samil_001.png></td>                                                    \n");
                htmlStr.Append("                                                                                                                                                                \n");
                htmlStr.Append("							</tr>                                                                                                                               \n");
                htmlStr.Append("						</table>                                                                                                                                \n");
                htmlStr.Append("				</span> <![endif]></td>                                                                                                                         \n");
                htmlStr.Append("				<td class=xl7127974 colspan=5                                                                                                                   \n");
                htmlStr.Append("					style='border-top: 2.0pt double windowtext; text-align: left; vertical-align: top; padding-bottom: 17px'>&nbsp;CO.,                         \n");
                htmlStr.Append("					Ltd</td>                                                                                                                                    \n");
                htmlStr.Append("				<td class=xl7027974                                                                                                                             \n");
                htmlStr.Append("					style='border-top: 2.0pt double windowtext; border-right: 2.0pt double windowtext;'>&nbsp;</td>                                             \n");
                htmlStr.Append("			</tr>                                                                                                                                               \n");
                htmlStr.Append("                                                                                                                                                                \n");
                htmlStr.Append("			<tr height=26 style='mso-height-source: userset; height: 17.25pt'>                                                                                     \n");
                htmlStr.Append("				<td height=26 class=xl8027974 colspan=8 style='height: 17.25pt'><span                                                                              \n");
                htmlStr.Append("					style='mso-spacerun: yes'>  </span>&#272;&#7883;a ch&#7881;                                                                                 \n");
                htmlStr.Append("					(Address): " + dt.Rows[0]["Seller_Address"] + "</td>                                                                                                      \n");
                htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("				<td class=xl6527974 colspan=4                                                                                                                   \n");
                htmlStr.Append("					style='border-right: 2.0pt double black'></td>                                                                                                         \n");
                htmlStr.Append("			</tr>                                                                                                                                               \n");
                htmlStr.Append("			<tr height=26 style='mso-height-source: userset; height: 17.25pt'>                                                                                     \n");
                htmlStr.Append("				<td height=26 class=xl8227974 colspan=7 style='height: 17.25pt'><span                                                                              \n");
                htmlStr.Append("					style='mso-spacerun: yes'>  </span>&#272;i&#7879;n                                                                                          \n");
                htmlStr.Append("					Tho&#7841;i(Tel):<span style='mso-spacerun: yes'>  </span><font                                                                             \n");
                htmlStr.Append("					class=font827974>" + dt.Rows[0]["Seller_Tel"] + ";<span                                                                                                  \n");
                htmlStr.Append("						style='mso-spacerun: yes'>  </span>Fax:<span                                                                                            \n");
                htmlStr.Append("						style='mso-spacerun: yes'>  </span>" + dt.Rows[0]["Seller_Fax"] + "</font></td>                                                                      \n");
                htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("				<td class=xl6627974 colspan=4                                                                                                                   \n");
                htmlStr.Append("					style='border-right: 2.0pt double black'>Ký hi&#7879;u/Serial:                                                                              \n");
                htmlStr.Append("					" + dt.Rows[0]["templateCode"] + "" + dt.Rows[0]["INVOICESERIALNO"] + "</td>                                                                                                                      \n");
                htmlStr.Append("			</tr>                                                                                                                                               \n");
                htmlStr.Append("			<tr height=28 style='mso-height-source: userset; height: 17.25pt'>                                                                                     \n");
                htmlStr.Append("				<td height=28 class=xl8227974 colspan=6 style='height: 17.25pt'><span                                                                              \n");
                htmlStr.Append("					style='mso-spacerun: yes'>  </span>" + dt.Rows[0]["Seller_Accountno"] + "</td>                                                                             \n");
                htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("				<td class=xl6627974 colspan=2>S&#7889;/No:</td>                                                                                                 \n");
                htmlStr.Append("				<td class=xl11527974 colspan=2>&nbsp;&nbsp;&nbsp;&nbsp;" + dt.Rows[0]["InvoiceNumber"] + "</td>                                                                  \n");
                htmlStr.Append("			</tr>                                                                                                                                               \n");
                htmlStr.Append("			<tr height=26 style='mso-height-source: userset; height: 17.25pt'>                                                                                     \n");
                htmlStr.Append("				<td height=26 class=xl8227974 colspan=7 style='height: 17.25pt'><span                                                                              \n");
                htmlStr.Append("					style='mso-spacerun: yes'>  </span>" + dt.Rows[0]["Seller_Bankname"] + "</td>                                                                                   \n");
                htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("				<td class=xl6627974 colspan=3>Ngày/Date:<span                                                                                                   \n");
                htmlStr.Append("					style='mso-spacerun: yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; " + dt.Rows[0]["INVOICEISSUEDDATE_DD"] + "/" + dt.Rows[0]["INVOICEISSUEDDATE_MM"] + "/" + dt.Rows[0]["INVOICEISSUEDDATE_YYYY"] + " </span></td>                                       \n");
                htmlStr.Append("				<td class=xl8127974></td>                                                                                                                       \n");
                htmlStr.Append("			</tr>                                                                                                                                               \n");
                htmlStr.Append("			<tr height=26 style='mso-height-source: userset; height: 17.25pt'>                                                                                     \n");
                htmlStr.Append("				<td height=26 class=xl8227974 colspan=2 style='height: 17.25pt'><span                                                                              \n");
                htmlStr.Append("					style='mso-spacerun: yes'>  </span>MST: <font class=font927974>" + dt.Rows[0]["Seller_Taxcode"] + "</font></td>                                          \n");
                htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("				<td class=xl8127974>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("			</tr>                                                                                                                                               \n");
                htmlStr.Append("			<tr height=50 style='mso-height-source: userset; height:27.5pt'>                                                                                     \n");
                htmlStr.Append("				<td colspan=15 height=50 class=xl15727974                                                                                                       \n");
                htmlStr.Append("					style='border-right: 2.0pt double black; height:27.5pt'><font                                                                                \n");
                htmlStr.Append("					class=font1227974>HÓA &#272;</font><font class=font1327974>&#416;</font><font                                                               \n");
                htmlStr.Append("					class=font1227974>N GIÁ TR&#7882; GIA T&#258;NG</font><font                                                                                 \n");
                htmlStr.Append("					class=font1427974> (</font><font class=font1527974>Value                                                                                    \n");
                htmlStr.Append("						added Invoice</font><font class=font1427974>)</font></td>                                                                               \n");
                htmlStr.Append("			</tr>                                                                                                                                               \n");
                htmlStr.Append("			<tr height=25 style='mso-height-source: userset; height: 20.00pt'>                                                                                     \n");
                htmlStr.Append("				<td colspan=15 height=25 class=xl16027974                                                                                                       \n");
                htmlStr.Append("					style='border-right: 2.0pt double black; height: 20.00pt'></td>                                                                                \n");
                htmlStr.Append("			</tr>                                                                                                                                               \n");
                htmlStr.Append("                                                                                                                                                                \n");
                htmlStr.Append("			<tr height=26 style='mso-height-source: userset; height: 16.25pt'>                                                                                     \n");
                htmlStr.Append("				<td colspan=4 height=26 class=xl9927974 style='height: 16.25pt'><span                                                                              \n");
                htmlStr.Append("					style='mso-spacerun: yes'>  </span>H&#7885; tên ng<font                                                                                     \n");
                htmlStr.Append("					class=font527974>&#432;</font><font class=font727974>&#7901;i                                                                               \n");
                htmlStr.Append("						mua</font><font class=font1127974>(Buyer's name)</font><font                                                                            \n");
                htmlStr.Append("					class=font727974>:</font></td>                                                                                                              \n");
                htmlStr.Append("				<td colspan=11 class=xl9227974                                                                                                                  \n");
                htmlStr.Append("					style='border-right: 2.0pt double black'>" + dt.Rows[0]["buyer"] + "</td>                                                                           \n");
                htmlStr.Append("			</tr>                                                                                                                                               \n");
                htmlStr.Append("			<tr height=26 style='mso-height-source: userset; height: 16.25pt'>                                                                                     \n");
                htmlStr.Append("				<td colspan=2 height=26 class=xl9927974 style='height: 16.25pt'><span                                                                              \n");
                htmlStr.Append("					style='mso-spacerun: yes'>  </span>&#272;&#7883;a ch&#7881; <font                                                                           \n");
                htmlStr.Append("					class=font1127974>(Address)</font><font class=font727974>:</font></td>                                                                      \n");
                htmlStr.Append("				<td colspan=13 class=xl9327974                                                                                                                  \n");
                htmlStr.Append("					style='border-right: 2.0pt double black; border-left: none'>&nbsp;" + dt.Rows[0]["Attribute_05"] + "</td>                                                    \n");
                htmlStr.Append("			</tr>                                                                                                                                               \n");
                htmlStr.Append("			<tr height=26 style='mso-height-source: userset; height: 16.25pt'>                                                                                     \n");
                htmlStr.Append("				<td colspan=3 height=26 class=xl9927974 style='height: 16.25pt'><span                                                                              \n");
                htmlStr.Append("					style='mso-spacerun: yes'>  </span>Tên đơn vị <font                                                                                         \n");
                htmlStr.Append("					class=font1127974>(Company)</font><font class=font727974>:</font></td>                                                                      \n");
                htmlStr.Append("				<td colspan=12 class=xl93279742                                                                                                                  \n");
                htmlStr.Append("					style='border-right: 2.0pt double black; border-left: none; border-top: 0.5pt soild black;'>&nbsp;" + dt.Rows[0]["buyerlegalname"] + "</td>                                                   \n");
                htmlStr.Append("			</tr>                                                                                                                                               \n");
                htmlStr.Append("			<tr height=26 style='mso-height-source: userset; height: 16.25pt'>																						 \n");
                htmlStr.Append("				<td colspan=2 height=26 class=xl9927974 style='height: 16.25pt'><span                                                                               \n");
                htmlStr.Append("					style='mso-spacerun: yes'>  </span>&#272;&#7883;a ch&#7881; <font                                                                            \n");
                htmlStr.Append("					class='font1127974'>(Address)</font><font class='font727974'>:</font></td>                                                                   \n");
                htmlStr.Append("				<td colspan=13 class=xl9327974                                                                                                                   \n");
                htmlStr.Append("					style='border-right: 2.0pt double black; border-left: none'>&nbsp;" + dt.Rows[0]["BuyerAddress"] + "</td>                                                       \n");
                htmlStr.Append("			</tr>                                                                                                                                                \n");
                htmlStr.Append("			<tr height=26 style='mso-height-source: userset; height: 16.25pt'>                                                                                      \n");
                htmlStr.Append("				<td colspan=2 height=26 class=xl9927974 style='height: 16.25pt'><span                                                                               \n");
                htmlStr.Append("					style='mso-spacerun: yes'>  </span>MST <font class='font1127974'>(Tax                                                                        \n");
                htmlStr.Append("						code)</font><font class='font727974'>:</font></td>                                                                                       \n");
                htmlStr.Append("				<td colspan=2 class=xl9327974 style='border-right: 2.0pt double white;'>&nbsp;" + dt.Rows[0]["BuyerTaxCode"] + "</td>                                                             \n");
                htmlStr.Append("				<td colspan=6 class=xl9927974 style='border-right:2.0pt double white;border-top: none;font-size: 12pt;'>Hình                                             \n");
                htmlStr.Append("					thức thanh toán <font class='font1127974'>(Term of payment):</font>                                                                          \n");
                htmlStr.Append("				</td>                                                                                                                                            \n");
                htmlStr.Append("				 <td colspan=5 class=xl9327974                                                                                                                   \n");
                htmlStr.Append("					style='border-right: 2.0pt double black'>&nbsp;&nbsp;&nbsp;&nbsp;" + dt.Rows[0]["PaymentMethodCK"] + "</td>                                                                   \n");
                htmlStr.Append("			</tr>                                                                                                                                                \n");
                htmlStr.Append("			<tr height=26 style='mso-height-source: userset; height: 16.25pt'>                                                                                     \n");
                htmlStr.Append("				<td colspan=4 height=26 class=xl9927974                                                                                                         \n");
                htmlStr.Append("					style='height: 16.25pt; border-bottom: none'><span                                                                                             \n");
                htmlStr.Append("					style='mso-spacerun: yes'>  </span>Giao hàng tại kho <font                                                                                  \n");
                htmlStr.Append("					class=font1127974>(Delivery place)</font><font class=font727974>:</font></td>                                                               \n");
                htmlStr.Append("				<td colspan=11 class=xl9327974>" + dt.Rows[0]["Attribute_01"] + "</td>                                                                                    \n");
                htmlStr.Append("			</tr>                                                                                                                                               \n");
                htmlStr.Append("			<tr height=26 style='mso-height-source: userset; height: 20.00pt'>                                                                                     \n");
                htmlStr.Append("				<td colspan=2 height=26 class=xl9927974 style='height: 20.00pt'><span                                                                              \n");
                htmlStr.Append("					style='mso-spacerun: yes'>  </span>S&#7889; order <font                                                                                     \n");
                htmlStr.Append("					class=font1127974>(No)</font><font class=font727974>:</font></td>                                                                           \n");
                htmlStr.Append("				<td colspan=13 rowspan=2 class=xl14727974                                                                                                       \n");
                htmlStr.Append("					style='border-right: 2.0pt double black; border-top: none; border-bottom: 1pt dotted black'>" + dt.Rows[0]["Attribute_04"] + "</td>                                \n");
                htmlStr.Append("			</tr>                                                                                                                                               \n");

                htmlStr.Append("			<tr height=26 style='mso-height-source: userset; height: 20.00pt'>                                                                                     \n");
                htmlStr.Append("				<td height=26 class=xl9927974 style='height: 20.00pt'>&nbsp;</td>                                                                                  \n");
                htmlStr.Append("				<td class=xl10027974>&nbsp;</td>                                                                                                                \n");
                htmlStr.Append("			</tr>                                                                                                                                               \n");
                
                htmlStr.Append("<tr height=26 style='mso-height-source: userset; height: 13pt'>                                                                                                 \n");
                htmlStr.Append("	<td colspan=3 height=26 class=xl9927974 style='height: 13pt'><span                                                                                          \n");
                htmlStr.Append("		style='mso-spacerun: yes'>  </span>Đơn vị tiền tệ <font class='font1127974'>(Currency)</font><font class='font727974'>:</font></td>                     \n");
                htmlStr.Append("	<td colspan=1 class=xl9327974 style='border-right: none'>&nbsp;"+dt.Rows[0]["CurrencyCodeUSD"]+"</td>                                                                        \n");
                htmlStr.Append("	<td colspan=5 class=xl99279741 style='border-right:2.0pt double white;'>Tỷ giá <font class='font1127974'>(Exchange Rate):</font>                            \n");
                htmlStr.Append("	</td>                                                                                                                                                       \n");
                htmlStr.Append("	 <td colspan=6 class=xl9327974                                                                                                                              \n");
                htmlStr.Append("		style='border-right: 2.0pt double black'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" + dt.Rows[0]["tr_rate_88"] + "</td>                                                            \n");
                htmlStr.Append("</tr>                                                                                                                                                           \n");


                htmlStr.Append("			<tr height=10 style='mso-height-source: userset; height: 0.5pt'>                                                                                    \n");
                htmlStr.Append("				<td height=10 class=xl82279741 style='height: 0.5pt'>&nbsp;</td>                                                                                \n");
                htmlStr.Append("				<td class=xl6627974 colspan=13></td>                                                                                                            \n");
                htmlStr.Append("				<td class=xl8127974></td>                                                                                                                       \n");
                htmlStr.Append("			</tr>                                                                                                                                               \n");
                htmlStr.Append("			<tr class=xl10127974 height=23 style='height: 16.25pt'>                                                                                                \n");
                htmlStr.Append("				<td height=23 class=xl10227974 width=59                                                                                                         \n");
                htmlStr.Append("					style='height: 16.25pt; width: 44pt'>STT</td>                                                                                                  \n");
                htmlStr.Append("				<td colspan=7 class=xl15127974 width=580                                                                                                        \n");
                htmlStr.Append("					style='border-left: none; width: 435pt'>Tên hàng hoá,                                                                                       \n");
                htmlStr.Append("					d&#7883;ch v&#7909;</td>                                                                                                                    \n");
                htmlStr.Append("				<td colspan=2 class=xl15127974 width=69                                                                                                         \n");
                htmlStr.Append("					style='border-left: none; width: 52pt'>&#272;VT</td>                                                                                        \n");
                htmlStr.Append("				<td colspan=2 class=xl15127974 width=97                                                                                                         \n");
                htmlStr.Append("					style='border-left: none; width: 73pt'>S&#7889; l<font                                                                                      \n");
                htmlStr.Append("					class=font627974>&#432;</font><font class=font927974>&#7907;ng</font></td>                                                                  \n");
                htmlStr.Append("				<td colspan=2 class=xl15127974 width=87                                                                                                         \n");
                htmlStr.Append("					style='border-left: none; width: 65pt'>&#272;<font                                                                                          \n");
                htmlStr.Append("					class=font627974>&#417;</font><font class=font927974>n                                                                                      \n");
                htmlStr.Append("						giá</font></td>                                                                                                                         \n");
                htmlStr.Append("				<td class=xl10327974 width=144                                                                                                                  \n");
                htmlStr.Append("					style='border-left: none; width: 108pt'>Thành ti&#7873;n</td>                                                                               \n");
                htmlStr.Append("			</tr>                                                                                                                                               \n");
                htmlStr.Append("			<tr class=xl6827974 height=23 style='height: 16.25pt'>                                                                                                 \n");
                htmlStr.Append("				<td height=23 class=xl10427974 width=59                                                                                                         \n");
                htmlStr.Append("					style='height: 16.25pt; width: 44pt'>No</td>                                                                                                   \n");
                htmlStr.Append("				<td colspan=7 class=xl14327974 width=580                                                                                                        \n");
                htmlStr.Append("					style='border-left: none; width: 435pt'>Commodity</td>                                                                                      \n");
                htmlStr.Append("				<td colspan=2 class=xl14327974 width=69                                                                                                         \n");
                htmlStr.Append("					style='border-left: none; width: 52pt'>Unit</td>                                                                                            \n");
                htmlStr.Append("				<td colspan=2 class=xl14327974 width=97                                                                                                         \n");
                htmlStr.Append("					style='border-left: none; width: 73pt'>Quantities</td>                                                                                      \n");
                htmlStr.Append("				<td colspan=2 class=xl14327974 width=87                                                                                                         \n");
                htmlStr.Append("					style='border-left: none; width: 65pt'>Unit price</td>                                                                                      \n");
                htmlStr.Append("				<td class=xl10527974 width=144                                                                                                                  \n");
                htmlStr.Append("					style='border-left: none; width: 108pt'>Amount</td>                                                                                         \n");
                htmlStr.Append("			</tr>                                                                                                                                               \n");
                htmlStr.Append("			<tr height=23 style='height: 16.25pt'>                                                                                                                 \n");
                htmlStr.Append("				<td height=23 class=xl8627974                                                                                                                   \n");
                htmlStr.Append("					style='height: 16.25pt; border-top: none'>1</td>                                                                                             \n");
                htmlStr.Append("				<td colspan=7 class=xl14527974 width=580                                                                                                        \n");
                htmlStr.Append("					style='border-left: none; width: 435pt'>2</td>                                                                                              \n");
                htmlStr.Append("				<td colspan=2 class=xl14627974 style='border-left: none'>3</td>                                                                                 \n");
                htmlStr.Append("				<td colspan=2 class=xl14627974 style='border-left: none'>4</td>                                                                                 \n");
                htmlStr.Append("				<td colspan=2 class=xl14627974 style='border-left: none'>5</td>                                                                                 \n");
                htmlStr.Append("				<td class=xl8727974 style='border-top: none; border-left: none'>6                                                                               \n");
                htmlStr.Append("					= 4 x 5</td>                                                                                                                                \n");
                htmlStr.Append("			</tr>                                                                                                                                               \n");
                htmlStr.Append("                                                                                                                                                                \n");

                v_rowHeight = "30.0pt"; //"26.5pt";
                v_rowHeightEmpty = "22.0pt";
                v_rowHeightNumber = 23.5;

                v_rowHeightLast = "30.0pt";// "23.5pt";
                v_rowHeightLastNumber = 23;// 23.5;
                v_rowHeightEmptyLast = "23.5pt"; //"23.5pt";


                for (int dtR = 0; dtR < page[k]; dtR++)
                {
                    if (!vlongItemName && dt_d.Rows[v_index]["ITEM_NAME"].ToString().Length >= 92)
                    {
                        v_rowHeight = "35.0pt"; //"26.5pt";    
                        v_rowHeightLast = "30.0pt"; //"27.5pt";
                        v_rowHeightLastNumber = 23.5;//27.5;
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

                        htmlStr.Append("			<tr height=54 style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                                     \n");
                        htmlStr.Append("				<td height=54 class=xl9827974                                                                                                                   \n");
                        htmlStr.Append("					style='height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "; border-bottom: 1pt dotted windowtext;'>" + dt_d.Rows[v_index][7] + "</td>                                                             \n");
                        htmlStr.Append("				<td colspan=7 class=xl13627974 width=580                                                                                                        \n");
                        htmlStr.Append("					style='border-right: 1pt solid black; border-bottom: 1pt dotted windowtext; border-left: none; width: 435pt'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString().Replace("&#xA;", "</br>&nbsp; ") + "</td>     \n");
                        htmlStr.Append("				<td colspan=2 class=xl13927974 width=69                                                                                                         \n");
                        htmlStr.Append("					style='border-left: none; width: 52pt; border-bottom: 1pt dotted windowtext;'>" + dt_d.Rows[v_index][1] + "</td>                                           \n");
                        htmlStr.Append("				<td colspan=2 class=xl14127974 width=97                                                                                                         \n");
                        htmlStr.Append("					style='width: 73pt; border-bottom: 1pt dotted windowtext;'>" + dt_d.Rows[v_index][2] + "&nbsp;</td>                                                        \n");
                        htmlStr.Append("				<td colspan=2 class=xl14127974 width=87                                                                                                         \n");
                        htmlStr.Append("					style='width: 65pt; border-bottom: 1pt dotted windowtext;'>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                                        \n");
                        htmlStr.Append("				<td class=xl9727974 style='border-bottom: 1pt dotted windowtext;'>" + dt_d.Rows[v_index][4] + "&nbsp;</td>                                                     \n");
                        htmlStr.Append("			</tr>                                                                                                                                               \n");

                    }
                    else if (dtR == page[k] - 1)//dong cuoi moi trang
                    {
                        if (k < v_countNumberOfPages - 1) //trang giua
                        {

                            htmlStr.Append("			<tr height=54 style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                                     \n");
                            htmlStr.Append("				<td height=54 class=xl98279741 style='height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "; border-top: none;'>" + dt_d.Rows[v_index][7] + "</td>                                                         \n");
                            htmlStr.Append("				<td colspan=7 class=xl136279741 width=580                                                                                                        \n");
                            htmlStr.Append("					style='border-right: 1pt solid black; border-left: none; width: 435pt;'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString().Replace("&#xA;", "</br>&nbsp; ") + "</td>                                            \n");
                            htmlStr.Append("				<td colspan=2 class=xl139279741 width=69                                                                                                         \n");
                            htmlStr.Append("					style='border-left: none; width: 52pt;border-bottom: 1pt soild windowtext;'>" + dt_d.Rows[v_index][1] + "&nbsp;</td>                                                                             \n");
                            htmlStr.Append("				<td colspan=2 class=xl141279741 width=97 style='width: 73pt;'>" + dt_d.Rows[v_index][2] + "&nbsp;</td>                                                            \n");
                            htmlStr.Append("				<td colspan=2 class=xl141279741 width=87 style='width: 65pt;'>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                                            \n");
                            htmlStr.Append("				<td class=xl97279741 style='border-top: none;'>" + dt_d.Rows[v_index][4] + "&nbsp;</td>                                                                           \n");
                            htmlStr.Append("			</tr>                                                                                                                                               \n");
                            htmlStr.Append("                                                                                                                                                                \n");

                        }
                        else // trang cuoi
                        {
                            if (dtR == rowsPerPage - 1) // du 11 dong
                            {
                                htmlStr.Append("			<tr height=54 style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                                     \n");
                                htmlStr.Append("				<td height=54 class=xl98279741 style='height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "; border-top: none;'>" + dt_d.Rows[v_index][7] + "</td>                                                         \n");
                                htmlStr.Append("				<td colspan=7 class=xl136279741 width=580                                                                                                        \n");
                                htmlStr.Append("					style='border-right: 1pt solid black; border-left: none; width: 435pt;'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString().Replace("&#xA;", "</br>&nbsp; ") + "</td>                                            \n");
                                htmlStr.Append("				<td colspan=2 class=xl139279741 width=69                                                                                                         \n");
                                htmlStr.Append("					style='border-left: none; width: 52pt;border-bottom: 1pt soild windowtext;'>" + dt_d.Rows[v_index][1] + "&nbsp;</td>                                                                             \n");
                                htmlStr.Append("				<td colspan=2 class=xl141279741 width=97 style='width: 73pt;'>" + dt_d.Rows[v_index][2] + "&nbsp;</td>                                                            \n");
                                htmlStr.Append("				<td colspan=2 class=xl141279741 width=87 style='width: 65pt;'>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                                            \n");
                                htmlStr.Append("				<td class=xl97279741 style='border-top: none;'>" + dt_d.Rows[v_index][4] + "&nbsp;</td>                                                                           \n");
                                htmlStr.Append("			</tr>                                                                                                                                               \n");
                                htmlStr.Append("                                                                                                                                                                \n");

                            }
                            else
                            {
                                htmlStr.Append("			<tr height=54 style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                                     \n");
                                htmlStr.Append("				<td height=54 class=xl98279741 style='height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "; border-top: none;'>" + dt_d.Rows[v_index][7] + "</td>                                                         \n");
                                htmlStr.Append("				<td colspan=7 class=xl136279741 width=580                                                                                                        \n");
                                htmlStr.Append("					style='border-right: 1pt solid black; border-left: none; width: 435pt;'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString().Replace("&#xA;", "</br>&nbsp; ") + "</td>                                            \n");
                                htmlStr.Append("				<td colspan=2 class=xl139279741 width=69                                                                                                         \n");
                                htmlStr.Append("					style='border-left: none; width: 52pt;border-bottom: 1pt soild windowtext;'>" + dt_d.Rows[v_index][1] + "&nbsp;</td>                                                                             \n");
                                htmlStr.Append("				<td colspan=2 class=xl141279741 width=97 style='width: 73pt;'>" + dt_d.Rows[v_index][2] + "&nbsp;</td>                                                            \n");
                                htmlStr.Append("				<td colspan=2 class=xl141279741 width=87 style='width: 65pt;'>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                                            \n");
                                htmlStr.Append("				<td class=xl97279741 style='border-top: none;'>" + dt_d.Rows[v_index][4] + "&nbsp;</td>                                                                           \n");
                                htmlStr.Append("			</tr>                                                                                                                                               \n");
                                htmlStr.Append("                                                                                                                                                                \n");

                            }

                        }
                    }
                    else
                    { // dong giua
                      // 

                        htmlStr.Append("			<tr height=54 style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                                     \n");
                        htmlStr.Append("				<td height=54 class=xl9827974                                                                                                                   \n");
                        htmlStr.Append("					style='height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "; border-top: none; border-bottom: 1pt dotted windowtext;'>" + dt_d.Rows[v_index][7] + "</td>                                           \n");
                        htmlStr.Append("				<td colspan=7 class=xl13627974 width=580                                                                                                        \n");
                        htmlStr.Append("					style='border-right: 1pt solid black; border-bottom: 1pt dotted windowtext; border-left: none; width: 435pt'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString().Replace("&#xA;", "</br>&nbsp; ") + "</td>     \n");
                        htmlStr.Append("				<td colspan=2 class=xl13927974 width=69                                                                                                         \n");
                        htmlStr.Append("					style='border-left: none; width: 52pt; border-bottom: 1pt dotted windowtext;'>" + dt_d.Rows[v_index][1] + "</td>                                           \n");
                        htmlStr.Append("				<td colspan=2 class=xl14127974 width=97                                                                                                         \n");
                        htmlStr.Append("					style='width: 73pt; border-bottom: 1pt dotted windowtext;'>" + dt_d.Rows[v_index][2] + "&nbsp;</td>                                                        \n");
                        htmlStr.Append("				<td colspan=2 class=xl14127974 width=87                                                                                                         \n");
                        htmlStr.Append("					style='width: 65pt; border-bottom: 1pt dotted windowtext;'>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                                        \n");
                        htmlStr.Append("				<td class=xl9727974                                                                                                                             \n");
                        htmlStr.Append("					style='border-top: none; border-bottom: 1pt dotted windowtext;'>" + dt_d.Rows[v_index][4] + "&nbsp;</td>                                                   \n");
                        htmlStr.Append("			</tr>                                                                                                                                               \n");
                        htmlStr.Append("                                                                                                                                                                \n");


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
                    v_spacePerPage = 10;
                }

                if (k == v_countNumberOfPages - 1 && page[k] < rowsPerPage) // Trang cuoi khong du dong
                {
                    v_rowHeightEmptyLast = Math.Round(v_totalHeightLastPage / (rowsPerPage - page[k]), 2).ToString() + "pt";
                    for (int i = 0; i < rowsPerPage - page[k]; i++)
                    {
                        if (i == (rowsPerPage - page[k] - 1))
                        {
                            htmlStr.Append("							<tr height=54 style='mso-height-source: userset; height: " + v_rowHeightEmptyLast + "'>                                                                                 \n");
                            htmlStr.Append("								<td height=54 class=xl9827974 style='height: " + v_rowHeightEmptyLast + "; border-top: none'></td>                                                                  \n");
                            htmlStr.Append("								<td colspan=7 class=xl13627974 width=580                                                                                                    \n");
                            htmlStr.Append("									style='border-right: .5pt solid black; border-left: none; width: 435pt'>&nbsp;</td>                                                     \n");
                            htmlStr.Append("								<td colspan=2 class=xl13927974 width=69                                                                                                     \n");
                            htmlStr.Append("									style='border-left: none; width: 52pt'>&nbsp;</td>                                                                                      \n");
                            htmlStr.Append("								<td colspan=2 class=xl14127974 width=97 style='width: 73pt'>&nbsp;</td>                                                                     \n");
                            htmlStr.Append("								<td colspan=2 class=xl14127974 width=87 style='width: 65pt'>&nbsp;</td>                                                                     \n");
                            htmlStr.Append("								<td class=xl9727974 style='border-top: none'>&nbsp;</td>                                                                                    \n");
                            htmlStr.Append("							</tr>                                                                                                                                           \n");

                        }
                        else
                        {
                            htmlStr.Append("							<tr height=54 style='mso-height-source: userset; height: " + v_rowHeightEmptyLast + "'>                                                                                 \n");
                            htmlStr.Append("								<td height=54 class=xl9827974                                                                                                               \n");
                            htmlStr.Append("									style='height: " + v_rowHeightEmptyLast + "; border-top: none; border-bottom: 1pt dotted windowtext;'></td>                                                    \n");
                            htmlStr.Append("								<td colspan=7 class=xl13627974 width=580                                                                                                    \n");
                            htmlStr.Append("									style='border-right: .5pt solid black; border-bottom: 1pt dotted windowtext; border-left: none; width: 435pt'></td>                    \n");
                            htmlStr.Append("								<td colspan=2 class=xl13927974 width=69                                                                                                     \n");
                            htmlStr.Append("									style='border-left: none; width: 52pt; border-bottom: 1pt dotted windowtext;'></td>                                                    \n");
                            htmlStr.Append("								<td colspan=2 class=xl14127974 width=97                                                                                                     \n");
                            htmlStr.Append("									style='width: 73pt; border-bottom: 1pt dotted windowtext;'>&nbsp;</td>                                                                 \n");
                            htmlStr.Append("								<td colspan=2 class=xl14127974 width=87                                                                                                     \n");
                            htmlStr.Append("									style='width: 65pt; border-bottom: 1pt dotted windowtext;'>&nbsp;</td>                                                                 \n");
                            htmlStr.Append("								<td class=xl9727974                                                                                                                         \n");
                            htmlStr.Append("									style='border-top: none; border-bottom: 1pt dotted windowtext;'>&nbsp;</td>                                                            \n");
                            htmlStr.Append("							</tr>                                                                                                                                           \n");
                        }
                    } // for

                }//Trang cuoi 11 dong

                if (k < v_countNumberOfPages - 1)
                {

                    htmlStr.Append("					<tr height=18 style='mso-height-source: userset;border-bottom: 2.0pt double windowtext; height: " + (v_spacePerPage).ToString() + "pt'>                                                                                       \n");
                    htmlStr.Append("						<td height=18 class=xl11727974 colspan=2 style='height: " + (v_spacePerPage).ToString() + "pt'></td>                                                                               \n");
                    htmlStr.Append("						<td class=xl6627974>&nbsp;</td>                                                                                                                     \n");
                    htmlStr.Append("						<td class=xl6627974>&nbsp;</td>                                                                                                                     \n");
                    htmlStr.Append("						<td class=xl6627974>&nbsp;</td>                                                                                                                     \n");
                    htmlStr.Append("						<td class=xl6627974>&nbsp;</td>                                                                                                                     \n");
                    htmlStr.Append("						<td class=xl6627974>&nbsp;</td>                                                                                                                     \n");
                    htmlStr.Append("						<td class=xl6627974>&nbsp;</td>                                                                                                                     \n");
                    htmlStr.Append("						<td class=xl6627974>&nbsp;</td>                                                                                                                     \n");
                    htmlStr.Append("						<td class=xl6627974>&nbsp;</td>                                                                                                                     \n");
                    htmlStr.Append("						<td class=xl6627974>&nbsp;</td>                                                                                                                     \n");
                    htmlStr.Append("						<td class=xl6627974>&nbsp;</td>                                                                                                                     \n");
                    htmlStr.Append("						<td class=xl6627974>&nbsp;</td>                                                                                                                     \n");
                    htmlStr.Append("						<td class=xl6627974>&nbsp;</td>                                                                                                                     \n");
                    htmlStr.Append("						<td class=xl8127974>&nbsp;</td>                                                                                                                     \n");
                    htmlStr.Append("					</tr>		                                                                                                                                            \n");

                
                    htmlStr.Append("	<table  border=0>                                                                                                                                                                                                 \n");
                    htmlStr.Append("		<tr height=5  style='height: 2pt'>                                                                                                                                                                \n");

                    htmlStr.Append("		</tr>      																																														\n");
                    htmlStr.Append("	</table>             																																										\n");

                }


            }// for k                                                                                                                             
            htmlStr.Append("			<tr height=28 style='mso-height-source: userset; height: 22.0pt'>                                                                                   \n");
            htmlStr.Append("				<td colspan=2 height=28 class=xl13227974 style='height: 22.0pt'><span                                                                           \n");
            htmlStr.Append("					style='mso-spacerun: yes'> </span>"+ lb_amount_trans+"<span                                                                       \n");
            htmlStr.Append("					style='mso-spacerun: yes'>                  </span></td>                                                                                    \n");
            htmlStr.Append("				<td colspan=2 class=xl13427974>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("				<td colspan=2 class=xl8827974 style='border-bottom: 1pt solid black'>&nbsp;" + amount_trans + "</td>                                                                                                  \n");
            htmlStr.Append("				<td class=xl8827974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl8827974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td colspan=6 class=xl8927974 style='border-right: 1pt solid black'>C&#7897;ng                                                                 \n");
            htmlStr.Append("					ti&#7873;n hàng <font class=font1127974>(Total):</font><font                                                                                \n");
            htmlStr.Append("					class=font727974><span style='mso-spacerun: yes'>  </span></font>                                                                           \n");
            htmlStr.Append("				</td>                                                                                                                                           \n");
            htmlStr.Append("				<td class=xl9527974 style='border-left: none'>" + amount_net + "&nbsp;</td>                                                                        \n");
            htmlStr.Append("			</tr>                                                                                                                                               \n");
            htmlStr.Append("			<tr height=28 style='mso-height-source: userset; height: 22.0pt'>                                                                                   \n");
            htmlStr.Append("				<td colspan=2 height=28 class=xl13227974 style='height: 22.0pt'><span                                                                           \n");
            htmlStr.Append("					style='mso-spacerun: yes'> </span>Thu&#7871; su&#7845;t <font                                                                               \n");
            htmlStr.Append("					class=font1127974>(Rate)</font><font class=font727974>:<span                                                                                \n");
            htmlStr.Append("						style='mso-spacerun: yes'>                                  </span></font></td>                                                         \n");
            htmlStr.Append("				<td class=xl8927974 style='border-top: none'>" + dt.Rows[0]["taxrate"] + "</td>                                                                              \n");
            htmlStr.Append("				<td class=xl8827974 style='border-top: none'>&nbsp;</td>                                                                                        \n");
            htmlStr.Append("				<td class=xl8827974 style='border-top: none'>&nbsp;</td>                                                                                        \n");
            htmlStr.Append("				<td class=xl9027974 style='border-top: none'>&nbsp;</td>                                                                                        \n");
            htmlStr.Append("				<td class=xl9027974 style='border-top: none'>&nbsp;</td>                                                                                        \n");
            htmlStr.Append("				<td class=xl9027974 style='border-top: none'>&nbsp;</td>                                                                                        \n");
            htmlStr.Append("				<td colspan=6 class=xl8927974 style='border-right: 1pt solid black'>Thu&#7871;                                                                 \n");
            htmlStr.Append("					GTGT <font class=font1127974>(VAT):</font><font                                                                                             \n");
            htmlStr.Append("					class=font727974><span style='mso-spacerun: yes'>  </span></font>                                                                           \n");
            htmlStr.Append("				</td>                                                                                                                                           \n");
            htmlStr.Append("				<td class=xl9527974 style='border-top: none; border-left: none'> " + amount_vat + " &nbsp;</td>                                                    \n");
            htmlStr.Append("			</tr>                                                                                                                                               \n");
            htmlStr.Append("			<tr height=28 style='mso-height-source: userset; height: 22.0pt'>                                                                                   \n");
            htmlStr.Append("				<td colspan=5 height=28 class=xl13227974 style='height: 22.0pt'>&nbsp;</td>                                                                     \n");
            htmlStr.Append("				<td class=xl8827974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl8827974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td colspan=7 class=xl8927974 style='border-right: 1pt solid black'><span                                                                      \n");
            htmlStr.Append("					style='mso-spacerun: yes'>        </span>T&#7893;ng s&#7889;                                                                                \n");
            htmlStr.Append("					ti&#7873;n thanh toán <font class=font1127974>(Invoice                                                                                      \n");
            htmlStr.Append("						total):&nbsp;</font></td>                                                                                                                     \n");
            htmlStr.Append("				<td class=xl9527974 style='border-left: none; font-weight:600;'>" + amount_total + "&nbsp;</td>                                                               \n");
            htmlStr.Append("			</tr>                                                                                                                                               \n");
            htmlStr.Append("			<tr height=30 style='mso-height-source: userset; height: 18pt'>                                                                                     \n");
            htmlStr.Append("				<td height=30 class=xl9127974 colspan=5 style='height: 18pt'><span                                                                              \n");
            htmlStr.Append("					style='mso-spacerun: yes'> </span>S&#7889; ti&#7873;n vi&#7871;t                                                                            \n");
            htmlStr.Append("					b&#7857;ng ch&#7919; <font class=font1127974>(Total amount                                                                                  \n");
            htmlStr.Append("						in word)</font><font class=font727974>:<span                                                                                            \n");
            htmlStr.Append("						style='mso-spacerun: yes'> </span></font></td>                                                                                          \n");
            htmlStr.Append("				<!--<td class=xl9027974 colspan=9 style='border-top: none'>&nbsp;</td>  -->                                                                     \n");
            htmlStr.Append("				<td colspan=10 class=xl9427974>&nbsp;&nbsp;" + read_prive + "</td>                                                                                      \n");
            htmlStr.Append("			</tr>                                                                                                                                               \n");
            htmlStr.Append("			<tr height=26 style='mso-height-source: userset; height: 14.25pt'>                                                                                  \n");
            htmlStr.Append("				<td colspan=15 height=26 class=xl12527974                                                                                                       \n");
            htmlStr.Append("					style='border-right: 2.0pt double black; height: 14.25pt'><span                                                                             \n");
            htmlStr.Append("					style='mso-spacerun: yes'>  </span></td>                                                                                                    \n");
            htmlStr.Append("			</tr>                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append("			<tr height=26 style='mso-height-source: userset; height: 17.5pt'>                                                                                     \n");
            htmlStr.Append("				<td colspan=5 height=26 class=xl12627974                                                                                                        \n");
            htmlStr.Append("					style='height: 17.5pt; vertical-align: center'>    </span>Người mua                                                                           \n");
            htmlStr.Append("					hàng <font class=font1027974>(Buyer ) </font></td>                                                                                          \n");
            htmlStr.Append("				<td colspan=3 class=xl12427974><span style='mso-spacerun: yes'></td>                                                                            \n");
            htmlStr.Append("				<td colspan=7 class=xl12427974                                                                                                                  \n");
            htmlStr.Append("					style='border-right: 2.0pt double black'><span                                                                                              \n");
            htmlStr.Append("					style='mso-spacerun: yes'>   </span>Người bán hàng<font                                                                                     \n");
            htmlStr.Append("					class=font1027974>(Seller)</font></td>                                                                                                      \n");
            htmlStr.Append("			</tr>                                                                                                                                               \n");
            htmlStr.Append("			<tr height=23 style='height: 17.5pt'>                                                                                                                 \n");
            htmlStr.Append("				<td colspan=5 height=23 class=xl6727974 style='height: 17.5pt'><span                                                                              \n");
            htmlStr.Append("					style='mso-spacerun: yes'>   </span>(Ký, ghi rõ h&#7885; tên)</td>                                                                          \n");
            htmlStr.Append("				<td colspan=3 class=xl6827974></td>                                                                                                             \n");
            htmlStr.Append("                                                                                                                                                                \n");
            htmlStr.Append("				<td colspan=7 class=xl6827974                                                                                                                   \n");
            htmlStr.Append("					style='border-right: 2.0pt double black'><span                                                                                              \n");
            htmlStr.Append("					style='mso-spacerun: yes'>  </span>(Ký, ghi rõ h&#7885; tên)</td>                                                                           \n");
            htmlStr.Append("			</tr>                                                                                                                                               \n");
            htmlStr.Append("			<tr height=23 style='height: 17.5pt'>                                                                                                                 \n");
            htmlStr.Append("				<td colspan=5 height=23 class=xl12827974 style='height: 17.5pt'><span                                                                             \n");
            htmlStr.Append("					style='mso-spacerun: yes'>    </span>(Signature, full name)</td>                                                                            \n");
            htmlStr.Append("				<td colspan=3 class=xl12927974></td>                                                                                                            \n");
            htmlStr.Append("				<td colspan=7 class=xl13027974                                                                                                                  \n");
            htmlStr.Append("					style='border-right: 2.0pt double black'><span                                                                                              \n");
            htmlStr.Append("					style='mso-spacerun: yes'>   </span>(Signature, full name)</td>                                                                             \n");
            htmlStr.Append("			</tr>                                                                                                                                               \n");
            htmlStr.Append("			<tr height=23 style='height: 10.4pt'>                                                                                                               \n");
            htmlStr.Append("				<td height=23 class=xl8227974 style='height: 10.4pt'>&nbsp;</td>                                                                                \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl8127974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("			</tr>                                                                                                                                               \n");
            htmlStr.Append("			<tr height=24 style='height: 13.0pt'>                                                                                                               \n");
            htmlStr.Append("				<td height=24 class=xl8227974 style='height: 13.0pt'>&nbsp;</td>                                                                                \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl10827974 colspan=3>Signature Valid</td>                                                                                             \n");

            if (dt.Rows[0]["sign_yn"].ToString() == "Y")
            {

                htmlStr.Append("				<td align=left valign=top><![if !vml]><span                                                                                                     \n");
                htmlStr.Append("					style='mso-ignore: vglayout; position: absolute; z-index: 2; margin-left: 18px; margin-top: 12px; width: 81px; height: 54px'><img           \n");
                htmlStr.Append("						width=81 height=54                                                                                                                      \n");
                htmlStr.Append("						src='D:\\webproject\\e-invoice-ws\\02.Web\\EInvoice\\img\\check_signed.png'></span>                                                              \n");
                htmlStr.Append("					<![endif]><span style='mso-ignore: vglayout2'>                                                                                              \n");
                htmlStr.Append("						<table cellpadding=0 cellspacing=0>                                                                                                     \n");
                htmlStr.Append("							<tr>                                                                                                                                \n");
                htmlStr.Append("								<td height=24 class=xl10927974 width=34                                                                                         \n");
                htmlStr.Append("									style='height: 18.0pt; width: 32.5pt'>&nbsp;</td>                                                                             \n");
                htmlStr.Append("							</tr>                                                                                                                               \n");
                htmlStr.Append("						</table>                                                                                                                                \n");
                htmlStr.Append("				</span></td>                                                                                                                                    \n");
            }
            else
            {
                htmlStr.Append("						<td height=18 class=xl10927974 width=39                                                                                                 \n");
                htmlStr.Append("							style='height: 14.1pt; width: 29pt'>&nbsp;</td>                                                                                     \n");
            }
            htmlStr.Append("				<td class=xl11027974>&nbsp;</td>                                                                                                                \n");
            htmlStr.Append("				<td class=xl11027974>&nbsp;</td>                                                                                                                \n");
            htmlStr.Append("				<td class=xl11127974>&nbsp;</td>                                                                                                                \n");
            htmlStr.Append("			</tr>                                                                                                                                               \n");
            htmlStr.Append("			<tr height=20 style='mso-height-source: userset; height: 13.0pt'>                                                                                   \n");
            htmlStr.Append("				<td height=20 class=xl8227974 style='height: 13.0pt'>&nbsp;</td>                                                                                \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td colspan=7 class=xl11827974                                                                                                                  \n");
            htmlStr.Append("					style='border-right: 2.0pt double black'><font                                                                                              \n");
            htmlStr.Append("					class=font1727974>&#272;&#432;&#7907;c ký b&#7903;i:</font><font                                                                            \n");
            htmlStr.Append("					class=font1627974> " + dt.Rows[0]["SignedBy"] + "</font></td>                                                                                            \n");
            htmlStr.Append("			</tr>                                                                                                                                               \n");
            htmlStr.Append("			<tr height=24 style='height: 13.0pt'>                                                                                                               \n");
            htmlStr.Append("				<td height=24 class=xl8227974 colspan=2 style='height: 13.0pt'></td>                                                                            \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl11227974 colspan=5>Ngày Ký: <font                                                                                                   \n");
            htmlStr.Append("					class=font1827974>" + dt.Rows[0]["SignedDate"] + "</font></td>                                                                                                \n");
            htmlStr.Append("				<td class=xl11327974>&nbsp;</td>                                                                                                                \n");
            htmlStr.Append("				<td class=xl11427974>&nbsp;</td>                                                                                                                \n");
            htmlStr.Append("			</tr>                                                                                                                                               \n");
            htmlStr.Append("			<tr height=26 style='mso-height-source: userset; height: 16.25pt'>                                                                                     \n");
            htmlStr.Append("				<td height=26 class=xl11627974 colspan=2 style='height: 16.25pt'>&nbsp;Mã CQT: " + dt.Rows[0]["cqt_mccqt_id"] + "</td>                                   \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td colspan=5 class=xl6827974>Mã nh&#7853;n hóa &#273;&#417;n:&nbsp;&nbsp;&nbsp;<font style='font-size: 13.75pt'>&nbsp;&nbsp;" + dt.Rows[0]["matracuu"] + "</font></td>                                                                                                       \n");
           
            htmlStr.Append("				<td colspan=2 class=xl6827974                                                                                                                   \n");
            htmlStr.Append("					style='border-right: 2.0pt double black'>&nbsp;</td>                                                                                        \n");
            htmlStr.Append("			</tr>                                                                                                                                               \n");
            htmlStr.Append("			<tr height=18 style='mso-height-source: userset; height: 13.5pt'>                                                                                   \n");
            htmlStr.Append("				<td height=18 class=xl11727974 colspan=2 style='height: 13.5pt'>&nbsp;Tra                                                                             \n");
            htmlStr.Append("					c&#7913;u t&#7841;i Website: <font class=font2127974><span                                                                                  \n");
            htmlStr.Append("						style='mso-spacerun: yes'> </span></font><font class=font2227974>" + dt.Rows[0]["WEBSITE_EI"] + "</font>                            \n");
            htmlStr.Append("				</td>                                                                                                                                           \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl6627974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("				<td class=xl8127974>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("			</tr>                                                                                                                                               \n");
            htmlStr.Append("			<tr height=28 style='mso-height-source: userset; height: 12.5pt'>                                                                                   \n");
            htmlStr.Append("				<td colspan=15 height=28 class=xl12127974                                                                                                       \n");
            htmlStr.Append("					style='border-right: 2.0pt double black; height: 12.5pt'>" + dt.Rows[0]["CONTRACT_INFO_EI"] + "                                                                                                       \n");
            htmlStr.Append("				</td>                                                                                                                                           \n");
            htmlStr.Append("			</tr>                                                                                                                                               \n");
            htmlStr.Append("			<![if supportMisalignedColumns]>                                                                                                                    \n");
            htmlStr.Append("			<tr height=0 style='display: none'>                                                                                                                 \n");
            htmlStr.Append("				<td width=59 style='width: 44pt'></td>                                                                                                          \n");
            htmlStr.Append("				<td width=135 style='width: 101pt'></td>                                                                                                        \n");
            htmlStr.Append("				<td width=115 style='width: 86pt'></td>                                                                                                         \n");
            htmlStr.Append("				<td width=40 style='width: 37.5pt'></td>                                                                                                          \n");
            htmlStr.Append("				<td width=29 style='width:27.5pt'></td>                                                                                                          \n");
            htmlStr.Append("				<td width=34 style='width: 32.5pt'></td>                                                                                                          \n");
            htmlStr.Append("				<td width=25 style='width: 19pt'></td>                                                                                                          \n");
            htmlStr.Append("				<td width=202 style='width: 151pt'></td>                                                                                                        \n");
            htmlStr.Append("				<td width=39 style='width: 29pt'></td>                                                                                                          \n");
            htmlStr.Append("				<td width=30 style='width: 28.75pt'></td>                                                                                                          \n");
            htmlStr.Append("				<td width=63 style='width: 47pt'></td>                                                                                                          \n");
            htmlStr.Append("				<td width=34 style='width: 32.5pt'></td>                                                                                                          \n");
            htmlStr.Append("				<td width=49 style='width: 46.25pt'></td>                                                                                                          \n");
            htmlStr.Append("				<td width=38 style='width: 28pt'></td>                                                                                                          \n");
            htmlStr.Append("				<td width=144 style='width: 108pt'></td>                                                                                                        \n");
            htmlStr.Append("			</tr>                                                                                                                                               \n");
            htmlStr.Append("			<![endif]>                                                                                                                                          \n");
            htmlStr.Append("		</table>                                                                                                                                                \n");
            htmlStr.Append("</body>                                                                                                                                                                                                 \n");
            htmlStr.Append("</html>               \n");

            

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