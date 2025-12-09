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
using Newtonsoft.Json;

namespace EInvoice.Company
{
    public class ParisBaguetteHCM_New
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

            int pos = 11, pos_lv = 20, v_count = 0, count_page = 0, count_page_v = 0, r = 0, x = 0;

            v_count = dt_d.Rows.Count;  //_Invoices.Inv[0].Invoice.Products.Product.Count();
            int[] page = new int[40] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            int[] page_new = new int[40] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };

            int v_index = -1, rowsPerPage = 20, row_cur = 0, row_count = 0, n_page = 0;

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
            
            //ESysLib.WriteLogError("CheckingTradeCodeCQT error :" + JsonConvert.SerializeObject(page) + "   -  v_count  " + v_count);



            string currency = "", read_prive = "", read_en = "", read_amount = "", amount_vat = "", amount_total = "", amount_trans = "", amount_net = "", lb_amount_trans = "";

            read_prive = dt.Rows[0]["amount_word_vie"].ToString();

            //read_en = dt.Rows[0]["amount_word_eng"].ToString();


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
                lb_amount_trans = dt.Rows[0]["EXCHANGERATE"].ToString();
                amount_trans = "Tổng cộng VND (<font class='font526793'>Amount VND</font>) : " + dt.Rows[0]["TOT_AMT_BK_93"].ToString();
                amount_total = dt.Rows[0]["tot_amt_tr_94"].ToString();
                amount_vat = dt.Rows[0]["VAT_TR_AMT_DIS_TR_91"].ToString();
                amount_net = dt.Rows[0]["NET_TR_AMT_DIS_TR_89"].ToString();

                // read_prive = Num2VNText(dt.Rows[0]["TotalAmountInWord"].ToString(), "USD");
            }

            //read_en = dt.Rows[0]["TotalAmountInWord"].ToString();
            int end = 0;
            int count = count_page_v + r;
            double height = 130;
            StringBuilder htmlStr = new StringBuilder("");
            string heigh = "", heigh_d = "";

            htmlStr.Append("<!DOCTYPE html PUBLIC '-//W3C//DTD HTML 4.01 Transitional//EN' 'http://www.w3.org/TR/html4/loose.dtd'>																									\n");
            htmlStr.Append("<html>                                                                                                                                                                                                 \n");
            htmlStr.Append("<head>                                                                                                                                                                                                 \n");
            htmlStr.Append("<meta http-equiv='Content-Type' content='text/html; charset= UTF-8'>                                                                                                                                    \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append("<script type='text/javascript'                                                                                                                                                                         \n");
            htmlStr.Append("	src='${ pageContext.request.contextPath}/system/syscommand.js'></script>                                                                                                                             \n");
            htmlStr.Append("<title>Report E-Invoice</title>                                                                                                                                                                        \n");
            htmlStr.Append("<!-- Normalize or reset CSS with your favorite library -->                                                                                                                                             \n");
            htmlStr.Append("<link rel='stylesheet'                                                                                                                                                                                 \n");
            //   htmlStr.Append("	href='https://cdnjs.cloudflare.com/ajax/libs/normalize/3.0.3/normalize.css'>                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append("<!-- Load paper.css for happy printing -->                                                                                                                                                             \n");
            htmlStr.Append("<link rel='stylesheet'                                                                                                                                                                                 \n");
            //   htmlStr.Append("	href='https://cdnjs.cloudflare.com/ajax/libs/paper-css/0.2.3/paper.css'>                                                                                                                            \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append("<!-- Set page size here: A5, A4 or A3 -->                                                                                                                                                              \n");
            htmlStr.Append("<!-- Set also 'landscape' if you need -->                                                                                                                                                              \n");
            htmlStr.Append("<style>                                                                                                                                                                                                \n");
            htmlStr.Append("@page {                                                                                                                                                                                                \n");
            htmlStr.Append("	size: A4                                                                                                                                                                                            \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("</style>                                                                                                                                                                                               \n");
            //   htmlStr.Append("<link href='https://fonts.googleapis.com/css?family=Tangerine:700'                                                                                                                                     \n");
            //   htmlStr.Append("	rel='stylesheet' type='text/css'>                                                                                                                                                                   \n");
            htmlStr.Append("<style>                                                                                                                                                                                                \n");
            htmlStr.Append("/*body   { font-family: serif }                                                                                                                                                                        \n");
            htmlStr.Append("    h1     { font-family: 'Tangerine', cursive; font-size: 48pt; line-height: 18mm}                                                                                                                    \n");
            htmlStr.Append("    h2, h3 { font-family: 'Tangerine', cursive; font-size: 16.8pt; line-height: 7mm }                                                                                                                    \n");
            htmlStr.Append("    h4     { font-size: 13pt; line-height: 1mm }                                                                                                                                                       \n");
            htmlStr.Append("    h2 + p { font-size: 18pt; line-height: 7mm }                                                                                                                                                       \n");
            htmlStr.Append("    h3 + p { font-size: 16.8pt; line-height: 7mm }                                                                                                                                                       \n");
            htmlStr.Append("    li     { font-size: 13.2pt; line-height: 5mm }                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append("    h1      { margin: 0 }                                                                                                                                                                              \n");
            htmlStr.Append("    h1 + ul { margin: 2mm 0 5mm }                                                                                                                                                                      \n");
            htmlStr.Append("    h2, h3  { margin: 0 3mm 3mm 0; float: left }                                                                                                                                                       \n");
            htmlStr.Append("    h2 + p,                                                                                                                                                                                            \n");
            htmlStr.Append("    h3 + p  { margin: 0 0 3mm 50mm }                                                                                                                                                                   \n");
            htmlStr.Append("    //h4      { margin: 1mm 0 0 2mm; border-bottom: 1px solid black }                                                                                                                                  \n");
            htmlStr.Append("    h4 + ul { margin: 5mm 0 0 50mm }                                                                                                                                                                   \n");
            htmlStr.Append("    article { border: 4px double black; padding: 5mm 10mm; border-radius: 3mm }*/                                                                                                                      \n");
            htmlStr.Append("body {                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: white;                                                                                                                                                                                        \n");
            htmlStr.Append("	font-size: 100%;                                                                                                                                                                                    \n");
            htmlStr.Append("	background-image: url('assets/Solution.jpg');                                                                                                                                                       \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append("h1 {                                                                                                                                                                                                   \n");
            htmlStr.Append("	color: #00FF00;                                                                                                                                                                                     \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append("p {                                                                                                                                                                                                    \n");
            htmlStr.Append("	color: rgb(0, 0, 255)                                                                                                                                                                               \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append("headline1 {                                                                                                                                                                                            \n");
            htmlStr.Append("	background-image: url(assets/Solution.jpg);                                                                                                                                                         \n");
            htmlStr.Append("	background-repeat: no-repeat;                                                                                                                                                                       \n");
            htmlStr.Append("	background-position: left top;                                                                                                                                                                      \n");
            htmlStr.Append("	padding-top: 68px;                                                                                                                                                                                  \n");
            htmlStr.Append("	margin-bottom: 50px;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append("headline2 {                                                                                                                                                                                            \n");
            htmlStr.Append("	background-image: url(images/newsletter_headline2.gif);                                                                                                                                             \n");
            htmlStr.Append("	background-repeat: no-repeat;                                                                                                                                                                       \n");
            htmlStr.Append("	background-position: left top;                                                                                                                                                                      \n");
            htmlStr.Append("	padding-top: 68px;                                                                                                                                                                                  \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append("<!--                                                                                                                                                                                                   \n");
            htmlStr.Append("table {                                                                                                                                                                                                \n");
            htmlStr.Append("	mso-displayed-decimal-separator: '\\.';                                                                                                                                                              \n");
            htmlStr.Append("	mso-displayed-thousand-separator: '\\, ';                                                                                                                                                             \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".font526793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".font626793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".font726793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".font826793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".font926793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	color: red;                                                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".font1026793 {                                                                                                                                                                                         \n");
            htmlStr.Append("	color: red;                                                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".font1126793 {                                                                                                                                                                                         \n");
            htmlStr.Append("	color: red;                                                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".font1226793 {                                                                                                                                                                                         \n");
            htmlStr.Append("	color: #993300;                                                                                                                                                                                     \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".font1326793 {                                                                                                                                                                                         \n");
            htmlStr.Append("	color: red;                                                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl6326793 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                             \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl6426793 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl6526793 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.4pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                             \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl6626793 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.4pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl6726793 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                             \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl6826793 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl6926793 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                             \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl7026793 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                             \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl7126793 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl7226793 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl7326793 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: 1pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl7426793 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: 1pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl7526793 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl7626793 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.4pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl7726793 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: 1pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl7826793 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: 1pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl7926793 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                             \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl8026793 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: '\\@';                                                                                                                                                                            \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                             \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl8126793 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl8226793 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl8326793 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl8426793 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: 1pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl8526793 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl8626793 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl8726793 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border: 1pt solid windowtext;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl8826793 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border: 1pt solid windowtext;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl8926793 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: 1pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: 1pt dotted windowtext;                                                                                                                                                              \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl9026793 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: 1pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: 1pt dotted windowtext;                                                                                                                                                              \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl9126793 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: 1pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: 1pt dotted windowtext;                                                                                                                                                              \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl9226793 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: 1pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl9326793 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl9426793 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl9526793 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                             \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl9626793 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                             \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl9726793 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: red;                                                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl9826793 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl9926793 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl10026793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: 1pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl10126793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl10226793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: red;                                                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl10326793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                             \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl10426793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: red;                                                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl10526793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: red;                                                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl10626793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: red;                                                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: 1pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl10726793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: red;                                                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl10826793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: red;                                                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl10926793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: red;                                                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: 1pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl11026793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: justify;                                                                                                                                                                            \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl11126793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl11226793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl11326793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl11426793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 11.8pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl11526793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl11626793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: '\\@';                                                                                                                                                                            \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                             \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl11726793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl11826793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl11926793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl12026793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl12126793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl12226793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: red;                                                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 21.6pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:1pt dotted windowtext;                                                                                                                                                              \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl12326793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;																																														\n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:1pt dotted windowtext;                                                                                                                                                              \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl12426793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom:1pt dotted windowtext;                                                                                                                                                              \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl12526793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: '\\#\\,\\#\\#0';                                                                                                                                                                     \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom:1pt dotted windowtext;                                                                                                                                                              \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl12626793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: '\\#\\,\\#\\#0';                                                                                                                                                                     \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl12726793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top:1pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:1pt dotted windowtext;                                                                                                                                                              \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("	text-transform: capitalize;                                                                                                                                                                         \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl12726793:first-letter {                                                                                                                                                                             \n");
            htmlStr.Append("	text-transform: capitalize;                                                                                                                                                                         \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl12826793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top:1pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:1pt dotted windowtext;                                                                                                                                                              \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl12926793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top:1pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom:1pt dotted windowtext;                                                                                                                                                              \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl13026793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: '\\#\\,\\#\\#0';                                                                                                                                                                     \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top:1pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom:1pt dotted windowtext;                                                                                                                                                              \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl13126793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: '\\#\\,\\#\\#0';                                                                                                                                                                     \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top:1pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom:1pt dotted windowtext;                                                                                                                                                              \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl13226793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top:1pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("	text-transform: capitalize;                                                                                                                                                                         \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl13326793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top:1pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl13426793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top:1pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl13526793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format:'_\\(* \\#\\,\\#\\#0\\.00_\\)\\;_\\(* \\\\(\\#\\,\\#\\#0\\.00\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                                                                           \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl13626793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl13726793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl13826793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl13926793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: '\\#\\,\\#\\#0\\;\\[Red\\]\\#\\,\\#\\#0';                                                                                                                                                   \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border:1pt solid windowtext;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl14026793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl14126793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: 0%;                                                                                                                                                                              \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl14226793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl14326793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl14426793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl14526793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                             \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl14626793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl14726793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl14826793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl14926793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl15026793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 21.6pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: justify;                                                                                                                                                                            \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl15126793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl15226793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl15326793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl15426793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl15526793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: #C00000;                                                                                                                                                                                     \n");
            htmlStr.Append("	font-size: 12p.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl15626793 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append(".xl136267931 {                                                                                                                                                                                          \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.2pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                  \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                      \n");
            htmlStr.Append(".xl13329656                                                                                                                                                                                              \n");
            htmlStr.Append("{                                                                                                                                                                                                        \n");
            htmlStr.Append("    padding: 0px;                                                                                                                                                                                        \n");
            htmlStr.Append("    mso-ignore:padding;                                                                                                                                                                                  \n");
            htmlStr.Append("    color: windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("    font-size:11.0pt;                                                                                                                                                                                    \n");
            htmlStr.Append("    font-weight:700;                                                                                                                                                                                     \n");
            htmlStr.Append("    font-style:normal;                                                                                                                                                                                   \n");
            htmlStr.Append("    text-decoration:none;                                                                                                                                                                                \n");
            htmlStr.Append("    font-family:'Times New Roman', serif;                                                                                                                                                                \n");
            htmlStr.Append("    mso-font-charset:0;                                                                                                                                                                                  \n");
            htmlStr.Append("    mso-number-format:General;                                                                                                                                                                           \n");
            htmlStr.Append("    text-align:center;                                                                                                                                                                                   \n");
            htmlStr.Append("    vertical-align:bottom;                                                                                                                                                                               \n");
            htmlStr.Append("    border-top:none;                                                                                                                                                                                     \n");
            htmlStr.Append("    border-right:none;                                                                                                                                                                                   \n");
            htmlStr.Append("    border-bottom:none;                                                                                                                                                                                  \n");
            htmlStr.Append("    border-left:.5pt solid windowtext;                                                                                                                                                                   \n");
            htmlStr.Append("    background: white;                                                                                                                                                                                   \n");
            htmlStr.Append("    mso-pattern:black none;                                                                                                                                                                              \n");
            htmlStr.Append("    white-space:nowrap;                                                                                                                                                                                  \n");
            htmlStr.Append("}                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl13429656                                                                                                                                                                                              \n");
            htmlStr.Append("{                                                                                                                                                                                                        \n");
            htmlStr.Append("    padding: 0px;                                                                                                                                                                                        \n");
            htmlStr.Append("    mso-ignore:padding;                                                                                                                                                                                  \n");
            htmlStr.Append("    color: windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("    font-size:11.0pt;                                                                                                                                                                                    \n");
            htmlStr.Append("    font-weight:700;                                                                                                                                                                                     \n");
            htmlStr.Append("    font-style:normal;                                                                                                                                                                                   \n");
            htmlStr.Append("    text-decoration:none;                                                                                                                                                                                \n");
            htmlStr.Append("    font-family:'Times New Roman', serif;                                                                                                                                                                \n");
            htmlStr.Append("    mso-font-charset:0;                                                                                                                                                                                  \n");
            htmlStr.Append("    mso-number-format:General;                                                                                                                                                                           \n");
            htmlStr.Append("    text-align:center;                                                                                                                                                                                   \n");
            htmlStr.Append("    vertical-align:bottom;                                                                                                                                                                               \n");
            htmlStr.Append("    background: white;                                                                                                                                                                                   \n");
            htmlStr.Append("    mso-pattern:black none;                                                                                                                                                                              \n");
            htmlStr.Append("    white-space:nowrap;                                                                                                                                                                                  \n");
            htmlStr.Append("}                                                                                                                                                                                                        \n");
            htmlStr.Append(".font829656                                                                                                                                                                                              \n");
            htmlStr.Append("{                                                                                                                                                                                                        \n");
            htmlStr.Append("    color: windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("    font-size:11.0pt;                                                                                                                                                                                    \n");
            htmlStr.Append("    font-weight:700;                                                                                                                                                                                     \n");
            htmlStr.Append("    font-style:italic;                                                                                                                                                                                   \n");
            htmlStr.Append("    text-decoration:none;                                                                                                                                                                                \n");
            htmlStr.Append("    font-family:'Times New Roman', serif;                                                                                                                                                                \n");
            htmlStr.Append("    mso-font-charset:0;                                                                                                                                                                                  \n");
            htmlStr.Append("}                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl6429656                                                                                                                                                                                               \n");
            htmlStr.Append("{                                                                                                                                                                                                        \n");
            htmlStr.Append("    padding: 0px;                                                                                                                                                                                        \n");
            htmlStr.Append("    mso-ignore:padding;                                                                                                                                                                                  \n");
            htmlStr.Append("    color: windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("    font-size:11.0pt;                                                                                                                                                                                    \n");
            htmlStr.Append("    font-weight:700;                                                                                                                                                                                     \n");
            htmlStr.Append("    font-style:normal;                                                                                                                                                                                   \n");
            htmlStr.Append("    text-decoration:none;                                                                                                                                                                                \n");
            htmlStr.Append("    font-family:'Times New Roman', serif;                                                                                                                                                                \n");
            htmlStr.Append("    mso-font-charset:0;                                                                                                                                                                                  \n");
            htmlStr.Append("    mso-number-format:General;                                                                                                                                                                           \n");
            htmlStr.Append("    text-align:general;                                                                                                                                                                                  \n");
            htmlStr.Append("    vertical-align:bottom;                                                                                                                                                                               \n");
            htmlStr.Append("    background: white;                                                                                                                                                                                   \n");
            htmlStr.Append("    mso-pattern:black none;                                                                                                                                                                              \n");
            htmlStr.Append("    white-space:nowrap;                                                                                                                                                                                  \n");
            htmlStr.Append("}                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl7929656                                                                                                                                                                                               \n");
            htmlStr.Append("{                                                                                                                                                                                                        \n");
            htmlStr.Append("    padding: 0px;                                                                                                                                                                                        \n");
            htmlStr.Append("    mso-ignore:padding;                                                                                                                                                                                  \n");
            htmlStr.Append("    color: windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("    font-size:10.0pt;                                                                                                                                                                                    \n");
            htmlStr.Append("    font-weight:400;                                                                                                                                                                                     \n");
            htmlStr.Append("    font-style:normal;                                                                                                                                                                                   \n");
            htmlStr.Append("    text-decoration:none;                                                                                                                                                                                \n");
            htmlStr.Append("    font-family:'Times New Roman', serif;                                                                                                                                                                \n");
            htmlStr.Append("    mso-font-charset:0;                                                                                                                                                                                  \n");
            htmlStr.Append("    mso-number-format:General;                                                                                                                                                                           \n");
            htmlStr.Append("    text-align:general;                                                                                                                                                                                  \n");
            htmlStr.Append("    vertical-align:bottom;                                                                                                                                                                               \n");
            htmlStr.Append("    border-top:none;                                                                                                                                                                                     \n");
            htmlStr.Append("    border-right:.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("    border-bottom:none;                                                                                                                                                                                  \n");
            htmlStr.Append("    border-left:none;                                                                                                                                                                                    \n");
            htmlStr.Append("    background: white;                                                                                                                                                                                   \n");
            htmlStr.Append("    mso-pattern:black none;                                                                                                                                                                              \n");
            htmlStr.Append("    white-space:nowrap;                                                                                                                                                                                  \n");
            htmlStr.Append("}                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl15129656                                                                                                                                                                                              \n");
            htmlStr.Append("{                                                                                                                                                                                                        \n");
            htmlStr.Append("    padding: 0px;                                                                                                                                                                                        \n");
            htmlStr.Append("    mso-ignore:padding;                                                                                                                                                                                  \n");
            htmlStr.Append("    color: windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("    font-size:10.0pt;                                                                                                                                                                                    \n");
            htmlStr.Append("    font-weight:400;                                                                                                                                                                                     \n");
            htmlStr.Append("    font-style:normal;                                                                                                                                                                                   \n");
            htmlStr.Append("    text-decoration:none;                                                                                                                                                                                \n");
            htmlStr.Append("    font-family:'Times New Roman', serif;                                                                                                                                                                \n");
            htmlStr.Append("    mso-font-charset:0;                                                                                                                                                                                  \n");
            htmlStr.Append("    mso-number-format:General;                                                                                                                                                                           \n");
            htmlStr.Append("    text-align:center;                                                                                                                                                                                   \n");
            htmlStr.Append("    vertical-align:bottom;                                                                                                                                                                               \n");
            htmlStr.Append("    border-top:none;                                                                                                                                                                                     \n");
            htmlStr.Append("    border-right:none;                                                                                                                                                                                   \n");
            htmlStr.Append("    border-bottom:none;                                                                                                                                                                                  \n");
            htmlStr.Append("    border-left:.5pt solid windowtext;                                                                                                                                                                   \n");
            htmlStr.Append("    background: white;                                                                                                                                                                                   \n");
            htmlStr.Append("    mso-pattern:black none;                                                                                                                                                                              \n");
            htmlStr.Append("    white-space:nowrap;                                                                                                                                                                                  \n");
            htmlStr.Append("}                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl10329656                                                                                                                                                                                              \n");
            htmlStr.Append("{                                                                                                                                                                                                        \n");
            htmlStr.Append("    padding: 0px;                                                                                                                                                                                        \n");
            htmlStr.Append("    mso-ignore:padding;                                                                                                                                                                                  \n");
            htmlStr.Append("    color: windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("    font-size:10.0pt;                                                                                                                                                                                    \n");
            htmlStr.Append("    font-weight:400;                                                                                                                                                                                     \n");
            htmlStr.Append("    font-style:normal;                                                                                                                                                                                   \n");
            htmlStr.Append("    text-decoration:none;                                                                                                                                                                                \n");
            htmlStr.Append("    font-family:'Times New Roman', serif;                                                                                                                                                                \n");
            htmlStr.Append("    mso-font-charset:0;                                                                                                                                                                                  \n");
            htmlStr.Append("    mso-number-format:General;                                                                                                                                                                           \n");
            htmlStr.Append("    text-align:center;                                                                                                                                                                                   \n");
            htmlStr.Append("    vertical-align:bottom;                                                                                                                                                                               \n");
            htmlStr.Append("    background: white;                                                                                                                                                                                   \n");
            htmlStr.Append("    mso-pattern:black none;                                                                                                                                                                              \n");
            htmlStr.Append("    white-space:nowrap;                                                                                                                                                                                  \n");
            htmlStr.Append("}                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl13529656                                                                                                                                                                                              \n");
            htmlStr.Append("{                                                                                                                                                                                                        \n");
            htmlStr.Append("    padding: 0px;                                                                                                                                                                                        \n");
            htmlStr.Append("    mso-ignore:padding;                                                                                                                                                                                  \n");
            htmlStr.Append("    color: windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("    font-size:10.0pt;                                                                                                                                                                                    \n");
            htmlStr.Append("    font-weight:400;                                                                                                                                                                                     \n");
            htmlStr.Append("    font-style:italic;                                                                                                                                                                                   \n");
            htmlStr.Append("    text-decoration:none;                                                                                                                                                                                \n");
            htmlStr.Append("    font-family:'Times New Roman', serif;                                                                                                                                                                \n");
            htmlStr.Append("    mso-font-charset:0;                                                                                                                                                                                  \n");
            htmlStr.Append("    mso-number-format:General;                                                                                                                                                                           \n");
            htmlStr.Append("    text-align:center;                                                                                                                                                                                   \n");
            htmlStr.Append("    vertical-align:bottom;                                                                                                                                                                               \n");
            htmlStr.Append("    border-top:none;                                                                                                                                                                                     \n");
            htmlStr.Append("    border-right:none;                                                                                                                                                                                   \n");
            htmlStr.Append("    border-bottom:none;                                                                                                                                                                                  \n");
            htmlStr.Append("    border-left:.5pt solid windowtext;                                                                                                                                                                   \n");
            htmlStr.Append("    background: white;                                                                                                                                                                                   \n");
            htmlStr.Append("    mso-pattern:black none;                                                                                                                                                                              \n");
            htmlStr.Append("    white-space:nowrap;                                                                                                                                                                                  \n");
            htmlStr.Append("}                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl13629656                                                                                                                                                                                              \n");
            htmlStr.Append("{                                                                                                                                                                                                        \n");
            htmlStr.Append("    padding: 0px;                                                                                                                                                                                        \n");
            htmlStr.Append("    mso-ignore:padding;                                                                                                                                                                                  \n");
            htmlStr.Append("    color: windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("    font-size:10.0pt;                                                                                                                                                                                    \n");
            htmlStr.Append("    font-weight:400;                                                                                                                                                                                     \n");
            htmlStr.Append("    font-style:italic;                                                                                                                                                                                   \n");
            htmlStr.Append("    text-decoration:none;                                                                                                                                                                                \n");
            htmlStr.Append("    font-family:'Times New Roman', serif;                                                                                                                                                                \n");
            htmlStr.Append("    mso-font-charset:0;                                                                                                                                                                                  \n");
            htmlStr.Append("    mso-number-format:General;                                                                                                                                                                           \n");
            htmlStr.Append("    text-align:center;                                                                                                                                                                                   \n");
            htmlStr.Append("    vertical-align:bottom;                                                                                                                                                                               \n");
            htmlStr.Append("    background: white;                                                                                                                                                                                   \n");
            htmlStr.Append("    mso-pattern:black none;                                                                                                                                                                              \n");
            htmlStr.Append("    white-space:nowrap;                                                                                                                                                                                  \n");
            htmlStr.Append("}                                                                                                                                                                                                        \n");
            htmlStr.Append("-->                                                                                                                                                                                                    \n");
            htmlStr.Append("</style>                                                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                                                       \n");
            htmlStr.Append("</head>                                                                                                                                                                                                \n");
            htmlStr.Append("<body class='A4'>                                                                                                                                                                                      \n");
            htmlStr.Append("	<table border=0 cellpadding=0 cellspacing=0 width=742 class=xl6326793                                                                                                                               \n");
            htmlStr.Append("		style='border-collapse: collapse; align=center table-layout: fixed; width: 666.4pt'>                                                                                                                           \n");
            htmlStr.Append("		<col class=xl6326793 width=19                                                                                                                                                                   \n");
            htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 682; width: 16.8pt'>                                                                                                                         \n");
            htmlStr.Append("		<col class=xl6326793 width=41                                                                                                                                                                   \n");
            htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 1450; width: 37.2pt'>                                                                                                                        \n");
            htmlStr.Append("		<col class=xl6326793 width=71                                                                                                                                                                   \n");
            htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 2531; width: 51.6pt'>                                                                                                                        \n");
            htmlStr.Append("		<col class=xl6326793 width=72                                                                                                                                                                   \n");
            htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 2560; width: 64.8pt'>                                                                                                                        \n");
            htmlStr.Append("		<col class=xl6326793 width=14                                                                                                                                                                   \n");
            htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 512; width: 53.2pt'>                                                                                                                         \n");
            htmlStr.Append("		<col class=xl6326793 width=22                                                                                                                                                                   \n");
            htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 796; width: 20.4pt'>                                                                                                                         \n");
            htmlStr.Append("		<col class=xl6326793 width=14                                                                                                                                                                   \n");
            htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 483; width: 12pt'>                                                                                                                         \n");
            htmlStr.Append("		<col class=xl6326793 width=25                                                                                                                                                                   \n");
            htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 881; width: 22.8pt'>                                                                                                                         \n");
            htmlStr.Append("		<col class=xl6326793 width=10                                                                                                                                                                   \n");
            htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 341; width:8.4pt'>                                                                                                                          \n");
            htmlStr.Append("		<col class=xl6326793 width=8                                                                                                                                                                    \n");
            htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 284; width:7.2pt'>                                                                                                                          \n");
            htmlStr.Append("		<col class=xl6326793 width=53                                                                                                                                                                   \n");
            htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 1877; width: 48pt'>                                                                                                                        \n");
            htmlStr.Append("		<col class=xl6326793 width=0                                                                                                                                                                    \n");
            htmlStr.Append("			style='display: none; mso-width-source: userset; mso-width-alt: 369'>                                                                                                                       \n");
            htmlStr.Append("		<col class=xl6326793 width=2                                                                                                                                                                    \n");
            htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 85; width:2.4pt'>                                                                                                                           \n");
            htmlStr.Append("		<col class=xl6326793 width=29                                                                                                                                                                   \n");
            htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 1024; width: 26.4pt'>                                                                                                                        \n");
            htmlStr.Append("		<col class=xl6326793 width=25                                                                                                                                                                   \n");
            htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 881; width: 22.8pt'>                                                                                                                         \n");
            htmlStr.Append("		<col class=xl6326793 width=18                                                                                                                                                                   \n");
            htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 654; width: 16.8pt'>                                                                                                                         \n");
            htmlStr.Append("		<col class=xl6326793 width=29                                                                                                                                                                   \n");
            htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 1024; width: 26.4pt'>                                                                                                                        \n");
            htmlStr.Append("		<col class=xl6326793 width=39                                                                                                                                                                   \n");
            htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 1393; width: 16.8pt'>                                                                                                                        \n");
            htmlStr.Append("		<col class=xl6326793 width=33                                                                                                                                                                   \n");
            htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 1166; width: 24pt'>                                                                                                                        \n");
            htmlStr.Append("		<col class=xl6326793 width=26                                                                                                                                                                   \n");
            htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 938; width: 18pt'>                                                                                                                         \n");
            htmlStr.Append("		<col class=xl6326793 width=37                                                                                                                                                                   \n");
            htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 1308; width: 30pt'>                                                                                                                        \n");
            htmlStr.Append("		<col class=xl6326793 width=25                                                                                                                                                                   \n");
            htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 881; width: 22.8pt'>                                                                                                                         \n");
            htmlStr.Append("		<col class=xl6326793 width=29                                                                                                                                                                   \n");
            htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 1024; width: 24pt'>                                                                                                                        \n");
            htmlStr.Append("		<col class=xl6326793 width=30                                                                                                                                                                   \n");
            htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 1052; width: 26.4pt'>                                                                                                                        \n");
            htmlStr.Append("		<col class=xl6326793 width=22                                                                                                                                                                   \n");
            htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 768; width: 25.2pt'>                                                                                                                         \n");
            htmlStr.Append("		<col class=xl6326793 width=28                                                                                                                                                                   \n");
            htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 995; width: 25.2pt'>                                                                                                                         \n");
            htmlStr.Append("		<col class=xl6326793 width=21                                                                                                                                                                   \n");
            htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 739; width: 16.8pt'>                                                                                                                         \n");



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

            double v_totalHeightPage = 518;//   540;
            bool v_break_page = false, v_type = false;
            int number_page = 0;
            for (int k = 0; k < v_countNumberOfPages; k++)
            {
                v_totalHeightPage = 565;// 540;
                number_page++;
                if (v_countNumberOfPages > 1)
                {
                    v_titlePageNumber = "Trang " + number_page.ToString() + "***";
                }

                if (k == v_countNumberOfPages - 1)
                {
                    rowsPerPage = pos;
                }
                else
                {
                    rowsPerPage = pos_lv;
                }
                htmlStr.Append("		<tr height=22 style='mso-height-source: userset; height: 19.8pt'>                                                                                                                               \n");
                htmlStr.Append("			<td height=22 class=xl7526793 width=19                                                                                                                                                      \n");
                htmlStr.Append("				style='height: 19.8pt; width: 16.8pt'>&nbsp;&nbsp;&nbsp;&nbsp;</td>                                                                                                                                         \n");
                htmlStr.Append("			<td colspan=24 class=xl11226793 width=674 style='width: 508pt'>" + dt.Rows[0]["Seller_Name"] + "<span                                                                                                           \n");
                htmlStr.Append("				style='mso-spacerun: yes'> </span></td>                                                                                                                                                 \n");
                htmlStr.Append("			<td class=xl7626793 width=28 style='width: 25.2pt'>&nbsp;</td>                                                                                                                                \n");
                htmlStr.Append("			<td class=xl8426793 width=21 style='width: 16pt'>&nbsp;</td>                                                                                                                                \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr height=46 style='mso-height-source: userset; height: 42.12pt'>                                                                                                                               \n");
                htmlStr.Append("			<td height=46 style='height: 42.12pt' align=left valign=top><span                                                                                                                            \n");
                htmlStr.Append("				style='mso-ignore: vglayout; position: absolute; z-index: 2; margin-left: 18px; margin-top: 1px; width: 362px; height: 65px'><img                                                       \n");
                htmlStr.Append("					width=362 height=65                                                                                                                                                                 \n");
                htmlStr.Append("					src='D:/webproject/e-invoice-ws/02.Web/EInvoice/img/PARIS_BAGUETTE_HCM_1.png'                                                                                                     \n");
                htmlStr.Append("					v:shapes='Picture_x0020_3'></span>                                                                                                                                                  \n");
                htmlStr.Append("			<![endif]><span style='mso-ignore: vglayout2'>                                                                                                                                              \n");
                htmlStr.Append("					<table cellpadding=0 cellspacing=0>                                                                                                                                                 \n");
                htmlStr.Append("						<tr>                                                                                                                                                                            \n");
                htmlStr.Append("							<td height=46 class=xl7126793 width=19                                                                                                                                      \n");
                htmlStr.Append("								style='height: 42.12pt; width: 16.8pt'>&nbsp;</td>                                                                                                                         \n");
                htmlStr.Append("						</tr>                                                                                                                                                                           \n");
                htmlStr.Append("					</table>                                                                                                                                                                            \n");
                htmlStr.Append("			</span></td>                                                                                                                                                                                \n");
                htmlStr.Append("			<td class=xl6926793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl7026793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl7026793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl7026793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl7026793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl7026793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl7026793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl7026793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl7026793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl7026793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl7026793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl7026793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl7026793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl7026793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl7026793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl7026793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl7026793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl7026793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl7026793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl7026793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl7026793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td align=left valign=top><span                                                                                                                                                             \n");
                htmlStr.Append("				style='mso-ignore: vglayout; position: absolute; z-index: 1; margin-left: 5px; margin-top: 0px; width:93px; height:84px'><img                                                         \n");
                htmlStr.Append("					width=93 height=84                                                                                                                                                                  \n");
                htmlStr.Append("					src='D:/webproject/e-invoice-ws/02.Web/EInvoice/img/PARIS_BAGUETTE_HCM_2.png'                                                                                                     \n");
                htmlStr.Append("					alt='logo vuong' v:shapes='Picture_x0020_2'></span>                                                                                                                                 \n");
                htmlStr.Append("			<![endif]><span style='mso-ignore: vglayout2'>                                                                                                                                              \n");
                htmlStr.Append("					<table cellpadding=0 cellspacing=0>                                                                                                                                                 \n");
                htmlStr.Append("						<tr>                                                                                                                                                                            \n");
                htmlStr.Append("							<td height=46 class=xl7026793 width=29                                                                                                                                      \n");
                htmlStr.Append("								style='height: 42.12pt; width: 26.4pt'></td>                                                                                                                               \n");
                htmlStr.Append("						</tr>                                                                                                                                                                           \n");
                htmlStr.Append("					</table>                                                                                                                                                                            \n");
                htmlStr.Append("			</span></td>                                                                                                                                                                                \n");
                htmlStr.Append("			<td class=xl7026793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl7026793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl6526793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl7726793>&nbsp;</td>                                                                                                                                                             \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr height=26 style='mso-height-source: userset; height: 19.8pt'>                                                                                                                               \n");
                htmlStr.Append("			<td height=26 class=xl7126793 style='height: 19.8pt'>&nbsp;</td>                                                                                                                            \n");
                htmlStr.Append("			<td colspan=24 class=xl11126793>Mã s&#7889; thu&#7871; (<font                                                                                                                               \n");
                htmlStr.Append("				class='font526793'>Tax code</font><font class='font626793'>):                                                                                                                           \n");
                htmlStr.Append("			</font><font class='font726793'>" + dt.Rows[0]["Seller_TaxCode"] + " </font></td>                                                                                                                             \n");
                htmlStr.Append("			<td class=xl6526793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl7726793>&nbsp;</td>                                                                                                                                                             \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr height=24 style='mso-height-source: userset; height: 19.8pt'>                                                                                                                               \n");
                htmlStr.Append("			<td height=24 class=xl7126793 style='height: 19.8pt'>&nbsp;</td>                                                                                                                            \n");
                htmlStr.Append("			<td colspan=26 class=xl11426793 width=674 style='width: 508pt;border-right:1pt solid windowtext;'>&#272;&#7883;a                                                                                                               \n");
                htmlStr.Append("				ch&#7881; (<font class='font526793'>Address</font><font                                                                                                                                 \n");
                htmlStr.Append("				class='font626793' style='width: 508pt;font-size:12.5pt'>): " + dt.Rows[0]["Seller_Address"] + " </font>                                                                                                                  \n");
                htmlStr.Append("			</td>                                                                                                                                                                                       \n");
                // htmlStr.Append("			<td class=xl6526793></td>                                                                                                                                                                   \n");
                // htmlStr.Append("			<td class=xl7726793>&nbsp;</td>                                                                                                                                                             \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr height=21 style='mso-height-source: userset; height: 19.8pt'>                                                                                                                               \n");
                htmlStr.Append("			<td height=21 class=xl7126793 style='height: 19.8pt'>&nbsp;</td>                                                                                                                            \n");
                htmlStr.Append("			<td colspan=24 class=xl11126793>&#272;i&#7879;n tho&#7841;i (<font                                                                                                                          \n");
                htmlStr.Append("				class='font526793'>Tel</font><font class='font626793'>)<span                                                                                                                            \n");
                htmlStr.Append("					style='mso-spacerun: yes'>   </span>: " + dt.Rows[0]["Seller_Tel"] + " <span                                                                                                                          \n");
                htmlStr.Append("					style='mso-spacerun: yes'>                      </span>Fax: " + dt.Rows[0]["Seller_Fax"] + " </font></td>                                                                                             \n");
                htmlStr.Append("			<td class=xl6526793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl7726793>&nbsp;</td>                                                                                                                                                             \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr height=22 style='mso-height-source: userset; height: 19.8pt'>                                                                                                                               \n");
                htmlStr.Append("			<td height=22 class=xl7126793 style='height: 19.8pt;'>&nbsp;</td>                                                                                                                           \n");
                htmlStr.Append("			<td colspan=24 class=xl11126793 style='border-bottom:.5pt solid black'>S&#7889; tài kho&#7843;n(<font                                                                                       \n");
                htmlStr.Append("				class='font526793'>Account No</font><font class='font626793'>)                                                                                                                          \n");
                htmlStr.Append("					: " + dt.Rows[0]["seller_accountno"] + " <span style='mso-spacerun: yes'>  </span>                                                                                                                      \n");
                htmlStr.Append("			</font></td>                                                                                                                                                                                \n");
                htmlStr.Append("			<td class=xl6526793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl7726793>&nbsp;</td>                                                                                                                                                             \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr height=26 style='mso-height-source: userset; height: 28.2pt'>                                                                                                                               \n");
                htmlStr.Append("			<td height=26 class=xl7526793 style='height: 28.2pt'>&nbsp;</td>                                                                                                                            \n");
                htmlStr.Append("			<td align=left valign=top><span                                                                                                                                                             \n");
                htmlStr.Append("				style='mso-ignore: vglayout; position: absolute; z-index: 1; margin-left: 5px; margin-top: 20px; width: 120px; height: 120px'><img                                                      \n");
                htmlStr.Append("					width=120 height=120                                                                                                                                                                \n");
                htmlStr.Append("					src='D:/webproject/e-invoice-ws/02.Web/EInvoice/img/QR_Code.png'                                                                                                                  \n");
                htmlStr.Append("					alt='logo vuong' v:shapes='Picture_x0020_2'></span>                                                                                                                                 \n");
                htmlStr.Append("			<![endif]><span style='mso-ignore: vglayout2'>                                                                                                                                              \n");
                htmlStr.Append("					<table cellpadding=0 cellspacing=0>                                                                                                                                                 \n");
                htmlStr.Append("						<tr>                                                                                                                                                                            \n");
                htmlStr.Append("							<td height=46 class=xl7626793 width=29                                                                                                                                      \n");
                htmlStr.Append("								style='height: 42.12pt; width: 26.4pt;border-top:none'></td>                                                                                                               \n");
                htmlStr.Append("						</tr>                                                                                                                                                                           \n");
                htmlStr.Append("					</table>                                                                                                                                                                            \n");
                htmlStr.Append("			</span></td>                                                                                                                                                                                \n");
                htmlStr.Append("			<td class=xl7626793>&nbsp;</td>                                                                                                                                                             \n");
                htmlStr.Append("			<td colspan=16 class=xl15026793><span style='mso-spacerun: yes'> </span>HÓA                                                                                                                 \n");
                htmlStr.Append("				&#272;&#416;N GIÁ TR&#7882; GIA T&#258;NG</td>                                                                                                                                               \n");
                htmlStr.Append("			<!-- <td class=xl15126793>&nbsp;</td> -->                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl9326793 colspan=4></td>                                                                                                                                                       \n");
                htmlStr.Append("			<td class=xl9326793 colspan=4 style='border-right:1pt solid black'> </td>                                                                                                       \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr height=26 style='mso-height-source: userset; height: 23.4pt'>                                                                                                                               \n");
                htmlStr.Append("			<td height=26 class=xl7126793 style='height: 23.4pt'>&nbsp;</td>                                                                                                                            \n");
                htmlStr.Append("			<td class=xl6526793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl6526793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td colspan=16 class=xl11026793>(VAT INVOICE)</td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl8626793 colspan=4><span style='mso-spacerun: yes'>                                                                                                                              \n");
                htmlStr.Append("			</span>Ký hi&#7879;u (<font class='font526793'>Serial</font><font                                                                                                                           \n");
                htmlStr.Append("				class='font626793'>):</font></td>                                                                                                                                                       \n");
                htmlStr.Append("			<td class=xl10126793 colspan=3>&nbsp;" + dt.Rows[0]["templateCode"] + "" + dt.Rows[0]["InvoiceSerialNo"] + " </td>                                                                                                                                     \n");
                htmlStr.Append("			<td class=xl7726793>&nbsp;</td>                                                                                                                                                             \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr height=20 style='mso-height-source: userset; height: 18.0pt'>                                                                                                                                           \n");
                htmlStr.Append("			<td height=20 class=xl7126793 style='height: 18.0pt'>&nbsp;</td>                                                                                                                                        \n");
                htmlStr.Append("			<td class=xl6526793></td>                                                                                                                                                                               \n");
                htmlStr.Append("			<td class=xl6526793></td>                                                                                                                                                                              \n");
                htmlStr.Append("			<td colspan=16 class=xl11126793 style='text-align: center'>"+ v_titlePageNumber  +"</td>                                                                                                                                     \n");
                htmlStr.Append("			<td class=xl8626793 colspan=3><span style='mso-spacerun: yes'>                                                                                                                                          \n");
                htmlStr.Append("			</span>S&#7889; (<font class='font526793'>No</font><font class='font626793'>)                                                                                                                           \n");
                htmlStr.Append("					:<span style='mso-spacerun: yes'>  </span>                                                                                                                                                      \n");
                htmlStr.Append("			</font></td>                                                                                                                                                                                            \n");
                htmlStr.Append("			<td class=xl10226793></td>                                                                                                                                                                              \n");
                htmlStr.Append("			<td class=xl10226793 colspan=3>" + dt.Rows[0]["InvoiceNumber"] + " </td>                                                                                                                                                      \n");
                htmlStr.Append("			<td class=xl7726793>&nbsp;</td>                                                                                                                                                                         \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                                       \n");
                htmlStr.Append("		<tr height=25 style='mso-height-source: userset; height: 22.68pt'>                                                                                                                               \n");
                htmlStr.Append("			<td height=25 class=xl7126793 style='height: 22.68pt'>&nbsp;</td>                                                                                                                            \n");
                htmlStr.Append("			<td class=xl6526793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl6526793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td colspan=16 class=xl11526793 style='text-align: right'>Ngày                                                                                                                              \n");
                htmlStr.Append("				(<font class='font526793'>Date</font><font class='font626793'>)</font>                                                                                                                  \n");
                htmlStr.Append("				 " + dt.Rows[0]["invoiceissueddate_dd"] + "  &nbsp;&nbsp; Tháng (<font class='font526793'>Month</font><font                                                                                                              \n");
                htmlStr.Append("				class='font626793'>)</font> " + dt.Rows[0]["invoiceissueddate_mm"] + "  &nbsp;&nbsp; N&#259;m (<font                                                                                                                     \n");
                htmlStr.Append("				class='font526793'>Year </font><font class='font626793'>)&nbsp;" + dt.Rows[0]["invoiceissueddate_yyyy"] + " &nbsp;&nbsp;                                                                                                 \n");
                htmlStr.Append("			</font>                                                                                                                                                                                     \n");
                htmlStr.Append("			</td>                                                                                                                                                                                       \n");
                htmlStr.Append("			<td class=xl6526793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl6526793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl6526793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl6526793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl6526793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl6526793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl6526793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl7726793>&nbsp;</td>                                                                                                                                                             \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr height=22 style='mso-height-source: userset; height: 20.16pt'>                                                                                                                               \n");
                htmlStr.Append("			<td height=22 class=xl7126793 style='height: 20.16pt'>&nbsp;</td>                                                                                                                            \n");
                htmlStr.Append("			<td colspan=5 class=xl11726793>H&#7885; tên ng&#432;&#7901;i mua                                                                                                                            \n");
                htmlStr.Append("				hàng (<font class='font526793'>Buyer</font><font class='font626793'>)                                                                                                                   \n");
                htmlStr.Append("					:</font>                                                                                                                                                                            \n");
                htmlStr.Append("			</td>                                                                                                                                                                                       \n");
                htmlStr.Append("			<td colspan=19 class=xl11726793>" + dt.Rows[0]["Buyer"] + "</td>                                                                                                                                         \n");
                htmlStr.Append("			<td class=xl8626793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl7726793>&nbsp;</td>                                                                                                                                                             \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr height=22 style='mso-height-source: userset; height: 20.7pt'>                                                                                                                              \n");
                htmlStr.Append("			<td height=22 class=xl7126793 style='height: 20.7pt'>&nbsp;</td>                                                                                                                           \n");
                htmlStr.Append("			<td colspan=4 class=xl11726793>Tên &#273;&#417;n v&#7883;<span                                                                                                                              \n");
                htmlStr.Append("				style='mso-spacerun: yes'>  </span>(<font class='font526793'>Company                                                                                                                    \n");
                htmlStr.Append("					name</font><font class='font626793'>) :</font></td>                                                                                                                                 \n");
                htmlStr.Append("			<td colspan=21 class=xl15626793>" + dt.Rows[0]["BuyerLegalName"] + " </td>                                                                                                                                         \n");
                htmlStr.Append("			<td class=xl7726793>&nbsp;</td>                                                                                                                                                             \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr height=22 style='mso-height-source: userset; height: 20.7pt'>                                                                                                                              \n");
                htmlStr.Append("			<td height=22 class=xl7126793 style='height: 20.7pt'>&nbsp;</td>                                                                                                                           \n");
                htmlStr.Append("			<td colspan=3 class=xl11726793>Mã s&#7889; thu&#7871;<span                                                                                                                                  \n");
                htmlStr.Append("				style='mso-spacerun: yes'>  </span>(<font class='font526793'>Tax                                                                                                                        \n");
                htmlStr.Append("					code</font><font class='font626793'>) :</font></td>                                                                                                                                 \n");
                htmlStr.Append("			<td colspan=21 class=xl11726793>" + dt.Rows[0]["BuyerTaxCode"] + " </td>                                                                                                                                           \n");
                htmlStr.Append("			<td class=xl8626793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl7726793>&nbsp;</td>                                                                                                                                                             \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr height=22 style='mso-height-source: userset; height: 20.7pt'>                                                                                                                              \n");
                htmlStr.Append("			<td height=22 class=xl7126793 style='height: 20.7pt'>&nbsp;</td>                                                                                                                           \n");
                htmlStr.Append("			<td colspan=3 class=xl11726793>&#272;&#7883;a ch&#7881; (<font                                                                                                                              \n");
                htmlStr.Append("				class='font526793'>Address)</font><font class='font626793'> :<span                                                                                                                      \n");
                htmlStr.Append("					style='mso-spacerun: yes'> </span></font></td>                                                                                                                                      \n");
                htmlStr.Append("			<td colspan=22 class=xl15626793>" + dt.Rows[0]["BuyerAddress"] + " </td>                                                                                                                                            \n");
                htmlStr.Append("			<td class=xl7726793>&nbsp;</td>                                                                                                                                                             \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr height=22 style='mso-height-source: userset; height: 19.8pt'>                                                                                                                               \n");
                htmlStr.Append("			<td height=22 class=xl7126793 style='height: 19.8pt'>&nbsp;</td>                                                                                                                            \n");
                htmlStr.Append("			<td class=xl8626793 colspan=8>Hình th&#7913;c thanh toán (<font                                                                                                                             \n");
                htmlStr.Append("				class='font526793'>Payment method</font><font class='font626793'>)                                                                                                                      \n");
                htmlStr.Append("					:<span style='mso-spacerun: yes'>   </span>                                                                                                                                         \n");
                htmlStr.Append("			</font></td>                                                                                                                                                                                \n");
                htmlStr.Append("			<td colspan=2 class=xl11826793>" + dt.Rows[0]["PaymentMethodCK"] + " </td>                                                                                                                                      \n");
                htmlStr.Append("			<td class=xl8626793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl8626793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl8626793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl8626793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td colspan=8 class=xl11726793>S&#7889; tài kho&#7843;n (<font                                                                                                                              \n");
                htmlStr.Append("				class='font526793'>Bank Account</font><font class='font626793'>):</font></td>                                                                                                           \n");
                htmlStr.Append("			<td colspan=2 class=xl11726793></td>                                                                                                                                                        \n");
                htmlStr.Append("			<td class=xl8626793></td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl7726793>&nbsp;</td>                                                                                                                                                             \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr class=xl6426793 height=37 style='height: 33.12pt'>                                                                                                                                           \n");
                htmlStr.Append("			<td height=37 class=xl7326793 style='border-top: none; height: 33.12pt'>&nbsp;</td>                                                                                                                            \n");
                htmlStr.Append("			<td class=xl8726793 width=41 style='border-left: none; width: 37.2pt'>STT                                                                                                                     \n");
                htmlStr.Append("				<font class='font526793'>(No.)</font>                                                                                                                                                   \n");
                htmlStr.Append("			</td>                                                                                                                                                                                       \n");
                htmlStr.Append("			<td colspan=11 class=xl11926793 width=291                                                                                                                                                   \n");
                htmlStr.Append("				style='border-right:1pt solid black; border-left: none; width: 222.8pt'>Tên                                                                                                             \n");
                htmlStr.Append("				hàng hóa, d&#7883;ch v&#7909;<font class='font526793'></br>(Description)</font>                                                                                                         \n");
                htmlStr.Append("			</td>                                                                                                                                                                                       \n");
                htmlStr.Append("			<td colspan=3 class=xl11926793 width=72                                                                                                                                                     \n");
                htmlStr.Append("				style='border-right:1pt solid black; border-left: none; width: 55pt'>&#272;vt</br>                                                                                                    \n");
                htmlStr.Append("			<font class='font526793'>(Unit)</font></td>                                                                                                                                                 \n");
                htmlStr.Append("			<td colspan=3 class=xl11926793 width=101                                                                                                                                                    \n");
                htmlStr.Append("				style='border-right:1pt solid black; border-left: none; width: 76pt'>S&#7889;                                                                                                         \n");
                htmlStr.Append("				l&#432;&#7907;ng <font class='font526793'>(Quantity)</font>                                                                                                                             \n");
                htmlStr.Append("			</td>                                                                                                                                                                                       \n");
                htmlStr.Append("			<td colspan=3 class=xl8726793 width=88                                                                                                                                                      \n");
                htmlStr.Append("				style='border-left: none; width: 67pt'>&#272;&#417;n giá <br>                                                                                                                           \n");
                htmlStr.Append("				<font class='font526793'>(Unit price)</font></td>                                                                                                                                       \n");
                htmlStr.Append("			<td colspan=4 class=xl8726793 width=109                                                                                                                                                     \n");
                htmlStr.Append("				style='border-left: none; width: 81pt'>Thành ti&#7873;n <font                                                                                                                           \n");
                htmlStr.Append("				class='font526793'>(Amount)</font></td>                                                                                                                                                 \n");
                htmlStr.Append("			<td class=xl7826793>&nbsp;</td>                                                                                                                                                             \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr class=xl6426793 height=18 style='height: 15.84pt'>                                                                                                                                           \n");
                htmlStr.Append("			<td height=18 class=xl7326793 style='height: 15.84pt'>&nbsp;</td>                                                                                                                            \n");
                htmlStr.Append("			<td class=xl8826793 width=41                                                                                                                                                                \n");
                htmlStr.Append("				style='border-top: none; border-left: none; width: 37.2pt'>1</td>                                                                                                                         \n");
                htmlStr.Append("			<td colspan=11 class=xl8826793 width=291                                                                                                                                                    \n");
                htmlStr.Append("				style='border-left: none; width: 222.8pt'>2</td>                                                                                                                                          \n");
                htmlStr.Append("			<td colspan=3 class=xl8826793 width=72                                                                                                                                                      \n");
                htmlStr.Append("				style='border-left: none; width: 55pt'>3</td>                                                                                                                                           \n");
                htmlStr.Append("			<td colspan=3 class=xl8826793 width=101                                                                                                                                                     \n");
                htmlStr.Append("				style='border-left: none; width: 76pt'>4</td>                                                                                                                                           \n");
                htmlStr.Append("			<td colspan=3 class=xl8826793 width=88                                                                                                                                                      \n");
                htmlStr.Append("				style='border-left: none; width: 67pt'>5</td>                                                                                                                                           \n");
                htmlStr.Append("			<td colspan=4 class=xl8826793 width=109                                                                                                                                                     \n");
                htmlStr.Append("				style='border-left: none; width: 81pt'>6 = 4 x 5</td>                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl7826793>&nbsp;</td>                                                                                                                                                             \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");

                v_rowHeight = "25.5pt"; //"26.5pt";
                v_rowHeightEmpty = "22.0pt";
                v_rowHeightNumber = 25.5;

                v_rowHeightLast = "25.5pt";// "23.5pt";
                v_rowHeightLastNumber = 25.5;// 23.5;
                v_rowHeightEmptyLast = "23.5pt"; //"23.5pt";

                int v_count_row_of_page = 0;
                for (int dtR = 0; dtR < page[k]; dtR++)
                {
                    v_count_row_of_page++;
                    if (dt_d.Rows[v_index]["ITEM_NAME"].ToString().Length >= 44)
                    {
                        v_count_row_of_page++;
                        //ESysLib.WriteLogError("SendPDF_R :  " + ex);
                        v_rowHeight = "28.5pt"; //"26.5pt";    
                        v_rowHeightLast = "28.5pt"; //"27.5pt";
                        v_rowHeightLastNumber = 32.5;//27.5;
                        v_rowHeightNumber = 32.5;//27.5;
                        v_rowHeightEmptyLast = "22.0pt"; //"23.0pt";
                        //vlongItemName = true;
                    }else
                    {
                        v_rowHeight = "25.5pt"; //"26.5pt";
                        v_rowHeightEmpty = "22.0pt";
                        v_rowHeightNumber = 25.5;
                        v_rowHeightLast = "25.5pt";// "23.5pt";
                        v_rowHeightLastNumber = 25.5;// 23.5;
                        v_rowHeightEmptyLast = "23.5pt"; //"23.5pt";
                    }

                    //ESysLib.WriteLogError("ParisBaguetteHCM_New error :" + k + " v_count_row_of_page " + v_count_row_of_page + " page[k]  " + page[k] + " dtR  " + dtR);

                    if (v_count_row_of_page > pos_lv && !v_break_page)
                    {
                        page[k + 1] = page[k + 1] + (page[k] - dtR);
                        break;
                    }
                    int end_row = v_count - v_index;

                    if (k == v_countNumberOfPages - 1 && !v_break_page)
                    {
                        //ESysLib.WriteLogError("ParisBaguetteHCM_New error 2:" + k + " v_count_row_of_page " + v_count_row_of_page + " page[k]  " + page[k] + " dtR  " + dtR);

                        //ESysLib.WriteLogError("ParisBaguetteHCM_New error 2: end_row " + end_row );

                        if (end_row > pos)
                        {
                            page[k] = 1;
                            page[k-1] = end_row - 1;
                            k = k - 1;
                            v_break_page = true;
                        }
                       
                        v_totalHeightPage = v_totalHeightLastPage + 310;
                      
                        //ESysLib.WriteLogError("ParisBaguetteHCM_New error 3: end_row " + JsonConvert.SerializeObject(page) + "  k" + k);
                    }

                    if (k == v_countNumberOfPages - 1 && !v_break_page && v_index == v_count - 1 && v_count_row_of_page > pos)
                    {
                       
                        page[k] = 1;
                        k = k - 1;
                        v_break_page = true;
                        v_type = false;
                        break;

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
                      
                        htmlStr.Append("		<tr class=xl6426793 height=26                                                                                                                                                                   \n");
                        htmlStr.Append("			style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                                                                                           \n");
                        htmlStr.Append("			<td height=26 class=xl7326793 style='height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>&nbsp;</td>                                                                                                                              \n");
                        htmlStr.Append("			<td class=xl9026793 width=41                                                                                                                                                                \n");
                        htmlStr.Append("				style='border-top: none; border-left: none; width: 37.2pt'>" + dt_d.Rows[v_index][7] + "</td>                                                                                                                 \n");
                        htmlStr.Append("			<td colspan=11 height=26 width=291                                                                                                                                                          \n");
                        htmlStr.Append("				style='border-right:1pt solid black; height: 19.8pt; width: 222.8pt'                                                                                                                    \n");
                        htmlStr.Append("				align=left valign=top><span                                                                                                                                                             \n");
                        htmlStr.Append("				style='mso-ignore: vglayout; position: absolute; z-index: 3; margin-left: 210px; margin-top: 8px; width: 262px; height: 276px'><img                                                     \n");
                        htmlStr.Append("					width=300 height=318                                                                                                                                                                \n");
                        htmlStr.Append("					src='D:/webproject/e-invoice-ws/02.Web/EInvoice/img/PARIS_BAGUETTE_HCM_0.png'                                                                                                     \n");
                        htmlStr.Append("					v:shapes='Picture_x0020_7'></span>                                                                                                                                                  \n");
                        htmlStr.Append("			<![endif]><span style='mso-ignore: vglayout2'>                                                                                                                                              \n");
                        htmlStr.Append("					<table cellpadding=0 cellspacing=0>                                                                                                                                                 \n");
                        htmlStr.Append("						<tr>                                                                                                                                                                            \n");
                        htmlStr.Append("							<td colspan=11 height=26 class=xl12726793 width=291                                                                                                                         \n");
                        htmlStr.Append("								style='border-right:1pt solid black; height: 19.8pt; border-left: none; border-right: none; border-top: none; border-bottom: none; width: 222.8pt'>&nbsp;" + dt_d.Rows[v_index][0] + "&nbsp;\n");
                        htmlStr.Append("							</td>                                                                                                                                                                       \n");
                        htmlStr.Append("						</tr>                                                                                                                                                                           \n");
                        htmlStr.Append("					</table>                                                                                                                                                                            \n");
                        htmlStr.Append("			</span></td>                                                                                                                                                                                \n");
                        htmlStr.Append("			<td colspan=3 class=xl9026793 width=72                                                                                                                                                      \n");
                        htmlStr.Append("				style='border-left: none; width: 55pt'>&nbsp;" + dt_d.Rows[v_index][1] + "&nbsp;                                                                                                                           \n");
                        htmlStr.Append("			</td>                                                                                                                                                                                       \n");
                        htmlStr.Append("			<td colspan=3 class=xl12526793 width=101                                                                                                                                                    \n");
                        htmlStr.Append("				style='border-left: none; width: 76pt'>&nbsp;" + dt_d.Rows[v_index][2] + "&nbsp;                                                                                                                           \n");
                        htmlStr.Append("			</td>                                                                                                                                                                                       \n");
                        htmlStr.Append("			<td colspan=3 class=xl12526793 width=88                                                                                                                                                     \n");
                        htmlStr.Append("				style='border-left: none; width: 67pt'>&nbsp;" + dt_d.Rows[v_index][3] + "&nbsp;                                                                                                                           \n");
                        htmlStr.Append("			</td>                                                                                                                                                                                       \n");
                        htmlStr.Append("			<td colspan=4 class=xl12626793 width=109                                                                                                                                                    \n");
                        htmlStr.Append("				style='border-left: none; width: 81pt'>&nbsp;" + dt_d.Rows[v_index][4] + "&nbsp;                                                                                                                           \n");
                        htmlStr.Append("			</td>                                                                                                                                                                                       \n");
                        htmlStr.Append("			<td class=xl7826793>&nbsp;</td>                                                                                                                                                             \n");
                        htmlStr.Append("		</tr>                                                                                                                                                                                           \n");


                    }
                    else if (dtR == page[k] - 1)//dong cuoi moi trang
                    {
                        if (k < v_countNumberOfPages - 1) //trang giua
                        {
                        
                            htmlStr.Append("		<tr height=27 style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                                                                                 \n");
                            htmlStr.Append("			<td height=27 class=xl7426793 style='height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>&nbsp;</td>                                                                                                                              \n");
                            htmlStr.Append("			<td class=xl9226793 style='border-left: none'>" + dt_d.Rows[v_index][7] + "&nbsp;</td>                                                                                                                         \n");
                            htmlStr.Append("			<td colspan=11 class=xl13226793                                                                                                                                                             \n");
                            htmlStr.Append("				style='border-right:1pt solid black; border-left: none'>&nbsp;" + dt_d.Rows[v_index][0] + "&nbsp;                                                                                                        \n");
                            htmlStr.Append("			</td>                                                                                                                                                                                       \n");
                            htmlStr.Append("			<td colspan=3 class=xl9226793 style='border-left: none'>" + dt_d.Rows[v_index][1] + "&nbsp;</td>                                                                                                               \n");
                            htmlStr.Append("			<td colspan=3 class=xl13526793 style='border-left: none'>" + dt_d.Rows[v_index][2] + "&nbsp;</td>                                                                                                              \n");
                            htmlStr.Append("			<td colspan=3 class=xl13526793 style='border-left: none'>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                                                                                              \n");
                            htmlStr.Append("			<td colspan=4 class=xl13526793 style='border-left: none'>" + dt_d.Rows[v_index][4] + "&nbsp;</td>                                                                                                              \n");
                            htmlStr.Append("			<td class=xl7726793>&nbsp;</td>                                                                                                                                                             \n");
                            htmlStr.Append("		</tr>                                                                                                                                                                                           \n");


                        }
                        else // trang cuoi
                        {
                            if (dtR == rowsPerPage - 1) // du 11 dong
                            {
                                htmlStr.Append("		<tr height=27 style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                                                                                 \n");
                                htmlStr.Append("			<td height=27 class=xl7426793 style='height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>&nbsp;</td>                                                                                                                              \n");
                                htmlStr.Append("			<td class=xl9226793 style='border-left: none'>" + dt_d.Rows[v_index][7] + "&nbsp;</td>                                                                                                                         \n");
                                htmlStr.Append("			<td colspan=11 class=xl13226793                                                                                                                                                             \n");
                                htmlStr.Append("				style='border-right:1pt solid black; border-left: none'>&nbsp;" + dt_d.Rows[v_index][0] + "&nbsp;                                                                                                        \n");
                                htmlStr.Append("			</td>                                                                                                                                                                                       \n");
                                htmlStr.Append("			<td colspan=3 class=xl9226793 style='border-left: none'>" + dt_d.Rows[v_index][1] + "&nbsp;</td>                                                                                                               \n");
                                htmlStr.Append("			<td colspan=3 class=xl13526793 style='border-left: none'>" + dt_d.Rows[v_index][2] + "&nbsp;</td>                                                                                                              \n");
                                htmlStr.Append("			<td colspan=3 class=xl13526793 style='border-left: none'>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                                                                                              \n");
                                htmlStr.Append("			<td colspan=4 class=xl13526793 style='border-left: none'>" + dt_d.Rows[v_index][4] + "&nbsp;</td>                                                                                                              \n");
                                htmlStr.Append("			<td class=xl7726793>&nbsp;</td>                                                                                                                                                             \n");
                                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");

                            }
                            else
                            {
                                //htmlStr.Append("		<tr height=27 style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                                                                                 \n");
                                //htmlStr.Append("			<td height=27 class=xl7426793 style='height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>&nbsp;</td>                                                                                                                              \n");
                                //htmlStr.Append("			<td class=xl9226793 style='border-left: none'>" + dt_d.Rows[v_index][7] + "&nbsp;</td>                                                                                                                         \n");
                                //htmlStr.Append("			<td colspan=11 class=xl13226793                                                                                                                                                             \n");
                                //htmlStr.Append("				style='border-right:1pt solid black; border-left: none'>&nbsp;" + dt_d.Rows[v_index][0] + "&nbsp;                                                                                                        \n");
                                //htmlStr.Append("			</td>                                                                                                                                                                                       \n");
                                //htmlStr.Append("			<td colspan=3 class=xl9226793 style='border-left: none'>" + dt_d.Rows[v_index][1] + "&nbsp;</td>                                                                                                               \n");
                                //htmlStr.Append("			<td colspan=3 class=xl13526793 style='border-left: none'>" + dt_d.Rows[v_index][2] + "&nbsp;</td>                                                                                                              \n");
                                //htmlStr.Append("			<td colspan=3 class=xl13526793 style='border-left: none'>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                                                                                              \n");
                                //htmlStr.Append("			<td colspan=4 class=xl13526793 style='border-left: none'>" + dt_d.Rows[v_index][4] + "&nbsp;KKK</td>                                                                                                              \n");
                                //htmlStr.Append("			<td class=xl7726793>&nbsp;</td>                                                                                                                                                             \n");
                                //htmlStr.Append("		</tr>                                                                                                                                                                                           \n");

                                htmlStr.Append("		<tr class=xl6426793 height=26                                                                                                                                                                   \n");
                                htmlStr.Append("			style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                                                                                           \n");
                                htmlStr.Append("			<td height=26 class=xl7326793 style='height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>&nbsp;</td>                                                                                                                              \n");
                                htmlStr.Append("			<td class=xl8926793 width=41                                                                                                                                                                \n");
                                htmlStr.Append("				style='border-top: none; border-left: none; width: 37.2pt'>" + dt_d.Rows[v_index][7] + "&nbsp;</td>                                                                                                           \n");
                                htmlStr.Append("			<td colspan=11 class=xl12726793                                                                                                                                                             \n");
                                htmlStr.Append("				style='border-right:1pt solid black; border-left: none'>&nbsp;" + dt_d.Rows[v_index][0] + "&nbsp;                                                                                                        \n");
                                htmlStr.Append("			</td>                                                                                                                                                                                       \n");
                                htmlStr.Append("			<td colspan=3 class=xl8926793 width=72                                                                                                                                                      \n");
                                htmlStr.Append("				style='border-left: none; width: 55pt'>" + dt_d.Rows[v_index][1] + "&nbsp;</td>                                                                                                                            \n");
                                htmlStr.Append("			<td colspan=3 class=xl13026793 width=101                                                                                                                                                    \n");
                                htmlStr.Append("				style='border-left: none; width: 76pt'>" + dt_d.Rows[v_index][2] + "&nbsp;</td>                                                                                                                            \n");
                                htmlStr.Append("			<td colspan=3 class=xl13026793 width=88                                                                                                                                                     \n");
                                htmlStr.Append("				style='border-left: none; width: 67pt'>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                                                                                                            \n");
                                htmlStr.Append("			<td colspan=4 class=xl13026793 width=109                                                                                                                                                    \n");
                                htmlStr.Append("				style='border-left: none; width: 81pt'>" + dt_d.Rows[v_index][4] + "&nbsp;</td>                                                                                                                            \n");
                                htmlStr.Append("			<td class=xl7826793>&nbsp;</td>                                                                                                                                                             \n");
                                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");


                            }
                        }
                    }
                    else
                    { // dong giua                                                                                                                                    
                        
                        htmlStr.Append("		<tr class=xl6426793 height=26                                                                                                                                                                   \n");
                        htmlStr.Append("			style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                                                                                           \n");
                        htmlStr.Append("			<td height=26 class=xl7326793 style='height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>&nbsp;</td>                                                                                                                              \n");
                        htmlStr.Append("			<td class=xl8926793 width=41                                                                                                                                                                \n");
                        htmlStr.Append("				style='border-top: none; border-left: none; width: 37.2pt'>" + dt_d.Rows[v_index][7] + "&nbsp;</td>                                                                                                           \n");
                        htmlStr.Append("			<td colspan=11 class=xl12726793                                                                                                                                                             \n");
                        htmlStr.Append("				style='border-right:1pt solid black; border-left: none'>&nbsp;" + dt_d.Rows[v_index][0] + "&nbsp;                                                                                                        \n");
                        htmlStr.Append("			</td>                                                                                                                                                                                       \n");
                        htmlStr.Append("			<td colspan=3 class=xl8926793 width=72                                                                                                                                                      \n");
                        htmlStr.Append("				style='border-left: none; width: 55pt'>" + dt_d.Rows[v_index][1] + "&nbsp;</td>                                                                                                                            \n");
                        htmlStr.Append("			<td colspan=3 class=xl13026793 width=101                                                                                                                                                    \n");
                        htmlStr.Append("				style='border-left: none; width: 76pt'>" + dt_d.Rows[v_index][2] + "&nbsp;</td>                                                                                                                            \n");
                        htmlStr.Append("			<td colspan=3 class=xl13026793 width=88                                                                                                                                                     \n");
                        htmlStr.Append("				style='border-left: none; width: 67pt'>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                                                                                                            \n");
                        htmlStr.Append("			<td colspan=4 class=xl13026793 width=109                                                                                                                                                    \n");
                        htmlStr.Append("				style='border-left: none; width: 81pt'>" + dt_d.Rows[v_index][4] + "&nbsp;</td>                                                                                                                            \n");
                        htmlStr.Append("			<td class=xl7826793>&nbsp;</td>                                                                                                                                                             \n");
                        htmlStr.Append("		</tr>                                                                                                                                                                                           \n");

                    }
                    v_index++;
                } //for dtR

                v_spacePerPage = 0;
                /*if (k < v_countNumberOfPages - 1 && page[k] < rowsPerPage)
                {
                    v_spacePerPage += v_totalHeightPage;
                }
                else if (k < v_countNumberOfPages - 1 && page[k] == rowsPerPage)
                {
                    v_spacePerPage = 50;
                }*/
                if (k < v_countNumberOfPages - 1)
                {
                    v_spacePerPage = v_totalHeightPage;
                }

                if (k == v_countNumberOfPages - 1 && page[k] < rowsPerPage) // Trang cuoi khong du dong
                {
                    v_rowHeightEmptyLast = Math.Round(v_totalHeightLastPage / (rowsPerPage - page[k]), 2).ToString() + "pt";
                    for (int i = 0; i < rowsPerPage - page[k]; i++)
                    {
                        if (i == (rowsPerPage - page[k] - 1))
                        {
                
                            htmlStr.Append("		<tr height=27 style='mso-height-source: userset; height: " + v_rowHeightEmptyLast + "'>                                                                                                                               \n");
                            htmlStr.Append("			<td height=27 class=xl7426793 style='height: " + v_rowHeightEmptyLast + "'>&nbsp;</td>                                                                                                                              \n");
                            htmlStr.Append("			<td class=xl9226793 style='border-left: none'>&nbsp;</td>                                                                                                                                   \n");
                            htmlStr.Append("			<td colspan=11 class=xl13226793                                                                                                                                                             \n");
                            htmlStr.Append("				style='border-right:1pt solid black; border-left: none'>&nbsp;</td>                                                                                                                   \n");
                            htmlStr.Append("			<td colspan=3 class=xl9226793 style='border-left: none'>&nbsp;</td>                                                                                                                         \n");
                            htmlStr.Append("			<td colspan=3 class=xl13526793 style='border-left: none'>&nbsp;</td>                                                                                                                        \n");
                            htmlStr.Append("			<td colspan=3 class=xl13526793 style='border-left: none'>&nbsp;</td>                                                                                                                        \n");
                            htmlStr.Append("			<td colspan=4 class=xl13526793 style='border-left: none'>&nbsp;</td>                                                                                                                        \n");
                            htmlStr.Append("			<td class=xl7726793>&nbsp;</td>                                                                                                                                                             \n");
                            htmlStr.Append("		</tr>                                                                                                                                                                                           \n");

                        }
                        else
                        {
        
                            htmlStr.Append("		<tr height=26 style='mso-height-source: userset; height: " + v_rowHeightEmptyLast + "'>                                                                                                                                 \n");
                            htmlStr.Append("			<td height=26 class=xl7426793 style='height: " + v_rowHeightEmptyLast + "'>&nbsp;</td>                                                                                                                              \n");
                            htmlStr.Append("			<td class=xl9126793 style='border-top: none; border-left: none'>&nbsp;</td>                                                                                                                 \n");
                            htmlStr.Append("			<td colspan=11 class=xl12726793                                                                                                                                                             \n");
                            htmlStr.Append("				style='border-right:1pt solid black; border-left: none'>&nbsp;</td>                                                                                                                   \n");
                            htmlStr.Append("			<td colspan=3 class=xl9126793 style='border-left: none'>&nbsp;</td>                                                                                                                         \n");
                            htmlStr.Append("			<td colspan=3 class=xl13126793 style='border-left: none'>&nbsp;</td>                                                                                                                        \n");
                            htmlStr.Append("			<td colspan=3 class=xl13126793 style='border-left: none'>&nbsp;</td>                                                                                                                        \n");
                            htmlStr.Append("			<td colspan=4 class=xl13126793 style='border-left: none'>&nbsp;</td>                                                                                                                        \n");
                            htmlStr.Append("			<td class=xl7726793>&nbsp;</td>                                                                                                                                                             \n");
                            htmlStr.Append("		</tr>                                                                                                                                                                                           \n");

                        }
                    } // for

                }//Trang cuoi 11 dong



                if (k < v_countNumberOfPages - 1)
                {
                  
                    htmlStr.Append("		<tr height=18 style='height: 13.2pt'>																					     \n");
                    htmlStr.Append("			<td colspan=27 height=18 class=xl14426793 width=742                                                                      \n");
                    htmlStr.Append("				style='border-right: .5pt solid black; height: 13.2pt; width: 559pt'></td>                                           \n");
                    htmlStr.Append("		</tr>                                                                                                                        \n");
                    htmlStr.Append("		<tr height=18 style='height: 13.2pt'>                                                                                        \n");
                    htmlStr.Append("			<td colspan=27 height=18 class=xl14726793                                                                                \n");
                    htmlStr.Append("				style='border-right: .5pt solid black; height: 13.2pt'></td>                                                         \n");
                    htmlStr.Append("		</tr>                                                                                                                        \n");
                    htmlStr.Append("		<tr height=10.05 style='mso-height-source: userset; height: " + (v_spacePerPage).ToString() + "pt'>                                                                                      \n");
                    htmlStr.Append("				<td colspan=27 height=10.05 class=xl10429612 style='height: " + (v_spacePerPage).ToString() + "pt; border-bottom:1.0pt solid ;'>&nbsp;</td>                                                                                  \n");
                    htmlStr.Append("		</tr>                                                                                                                                                   \n");
                    htmlStr.Append("	</table>                                                                                                                                                                                            \n");
                    htmlStr.Append("	<table border=0 cellpadding=0 cellspacing=0 width=742 class=xl6326793                                                                                                                               \n");
                    htmlStr.Append("		style='border-collapse: collapse; align=center table-layout: fixed; width: 666.4pt'>                                                                                                                           \n");
                    htmlStr.Append("		<col class=xl6326793 width=19                                                                                                                                                                   \n");
                    htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 682; width: 16.8pt'>                                                                                                                         \n");
                    htmlStr.Append("		<col class=xl6326793 width=41                                                                                                                                                                   \n");
                    htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 1450; width: 37.2pt'>                                                                                                                        \n");
                    htmlStr.Append("		<col class=xl6326793 width=71                                                                                                                                                                   \n");
                    htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 2531; width: 51.6pt'>                                                                                                                        \n");
                    htmlStr.Append("		<col class=xl6326793 width=72                                                                                                                                                                   \n");
                    htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 2560; width: 64.8pt'>                                                                                                                        \n");
                    htmlStr.Append("		<col class=xl6326793 width=14                                                                                                                                                                   \n");
                    htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 512; width: 53.2pt'>                                                                                                                         \n");
                    htmlStr.Append("		<col class=xl6326793 width=22                                                                                                                                                                   \n");
                    htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 796; width: 20.4pt'>                                                                                                                         \n");
                    htmlStr.Append("		<col class=xl6326793 width=14                                                                                                                                                                   \n");
                    htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 483; width: 12pt'>                                                                                                                         \n");
                    htmlStr.Append("		<col class=xl6326793 width=25                                                                                                                                                                   \n");
                    htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 881; width: 22.8pt'>                                                                                                                         \n");
                    htmlStr.Append("		<col class=xl6326793 width=10                                                                                                                                                                   \n");
                    htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 341; width:8.4pt'>                                                                                                                          \n");
                    htmlStr.Append("		<col class=xl6326793 width=8                                                                                                                                                                    \n");
                    htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 284; width:7.2pt'>                                                                                                                          \n");
                    htmlStr.Append("		<col class=xl6326793 width=53                                                                                                                                                                   \n");
                    htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 1877; width: 48pt'>                                                                                                                        \n");
                    htmlStr.Append("		<col class=xl6326793 width=0                                                                                                                                                                    \n");
                    htmlStr.Append("			style='display: none; mso-width-source: userset; mso-width-alt: 369'>                                                                                                                       \n");
                    htmlStr.Append("		<col class=xl6326793 width=2                                                                                                                                                                    \n");
                    htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 85; width:2.4pt'>                                                                                                                           \n");
                    htmlStr.Append("		<col class=xl6326793 width=29                                                                                                                                                                   \n");
                    htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 1024; width: 26.4pt'>                                                                                                                        \n");
                    htmlStr.Append("		<col class=xl6326793 width=25                                                                                                                                                                   \n");
                    htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 881; width: 22.8pt'>                                                                                                                         \n");
                    htmlStr.Append("		<col class=xl6326793 width=18                                                                                                                                                                   \n");
                    htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 654; width: 16.8pt'>                                                                                                                         \n");
                    htmlStr.Append("		<col class=xl6326793 width=29                                                                                                                                                                   \n");
                    htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 1024; width: 26.4pt'>                                                                                                                        \n");
                    htmlStr.Append("		<col class=xl6326793 width=39                                                                                                                                                                   \n");
                    htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 1393; width: 16.8pt'>                                                                                                                        \n");
                    htmlStr.Append("		<col class=xl6326793 width=33                                                                                                                                                                   \n");
                    htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 1166; width: 24pt'>                                                                                                                        \n");
                    htmlStr.Append("		<col class=xl6326793 width=26                                                                                                                                                                   \n");
                    htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 938; width: 18pt'>                                                                                                                         \n");
                    htmlStr.Append("		<col class=xl6326793 width=37                                                                                                                                                                   \n");
                    htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 1308; width: 30pt'>                                                                                                                        \n");
                    htmlStr.Append("		<col class=xl6326793 width=25                                                                                                                                                                   \n");
                    htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 881; width: 22.8pt'>                                                                                                                         \n");
                    htmlStr.Append("		<col class=xl6326793 width=29                                                                                                                                                                   \n");
                    htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 1024; width: 24pt'>                                                                                                                        \n");
                    htmlStr.Append("		<col class=xl6326793 width=30                                                                                                                                                                   \n");
                    htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 1052; width: 26.4pt'>                                                                                                                        \n");
                    htmlStr.Append("		<col class=xl6326793 width=22                                                                                                                                                                   \n");
                    htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 768; width: 25.2pt'>                                                                                                                         \n");
                    htmlStr.Append("		<col class=xl6326793 width=28                                                                                                                                                                   \n");
                    htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 995; width: 25.2pt'>                                                                                                                         \n");
                    htmlStr.Append("		<col class=xl6326793 width=21                                                                                                                                                                   \n");
                    htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 739; width: 16.8pt'>                                                                                                                         \n");

                }

            }// for k                                                                                                                             
            htmlStr.Append("		<tr class=xl6426793 height=26                                                                                                                                                                   \n");
            htmlStr.Append("			style='mso-height-source: userset; height: 20.1pt'>                                                                                                                                         \n");
            htmlStr.Append("			<td height=26 class=xl7326793 style='height: 20.1pt'>&nbsp;</td>                                                                                                                            \n");
            htmlStr.Append("			<td class=xl8126793 style='border-left: none' colspan=11>&nbsp;" + lb_amount_trans + "</td>                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl8526793>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("			<td colspan=9 class=xl13626793 style='border-right:1pt solid black'>&nbsp;C&#7897;ng                                                                                                      \n");
            htmlStr.Append("				ti&#7873;n hàng (<font class='font826793'>Total amount</font><font                                                                                                                      \n");
            htmlStr.Append("				class='font726793'>)<span style='mso-spacerun: yes'>                                                                                                                                    \n");
            htmlStr.Append("				</span>:<span style='mso-spacerun: yes'> </span></font>                                                                                                                                 \n");
            htmlStr.Append("			</td>                                                                                                                                                                                       \n");
            htmlStr.Append("			<td colspan=4 class=xl13926793 style='border-left: none'>" + dt.Rows[0]["netamount_display"] + currency + " &nbsp;</td>                                                                                                             \n");
            htmlStr.Append("			<td class=xl7826793>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
            htmlStr.Append("		<tr class=xl6426793 height=26                                                                                                                                                                   \n");
            htmlStr.Append("			style='mso-height-source: userset; height: 20.1pt'>                                                                                                                                         \n");
            htmlStr.Append("			<td height=26 class=xl7326793 style='height: 20.1pt'>&nbsp;</td>                                                                                                                            \n");
            htmlStr.Append("			<td colspan=4 class=xl9426793 style='border-left: none'><span                                                                                                                               \n");
            htmlStr.Append("				style='mso-spacerun: yes'> </span>Thu&#7871; su&#7845;t GTGT (<font                                                                                                                     \n");
            htmlStr.Append("				class='font526793'>VAT rate</font><font class='font626793'>):</font></td>                                                                                                               \n");
            htmlStr.Append("			<td colspan=3 class=xl14126793></td>                                                                                                                                                        \n");
            htmlStr.Append("			<td colspan=2 class=xl14226793>" + dt.Rows[0]["TaxRate"] + " &nbsp;</td>                                                                                                                                      \n");
            htmlStr.Append("			<td class=xl8226793 style='border-top: none'>&nbsp;</td>                                                                                                                                    \n");
            htmlStr.Append("			<td class=xl8226793 style='border-top: none'>&nbsp;</td>                                                                                                                                    \n");
            htmlStr.Append("			<td class=xl8226793 style='border-top: none'>&nbsp;</td>                                                                                                                                    \n");
            htmlStr.Append("			<td colspan=9 class=xl13626793 style='border-right:1pt solid black'>&nbsp;Ti&#7873;n                                                                                                      \n");
            htmlStr.Append("				thu&#7871; GTGT (<font class='font826793'>VAT</font><font                                                                                                                               \n");
            htmlStr.Append("				class='font726793'>)<span style='mso-spacerun: yes'>                                                                                                                                    \n");
            htmlStr.Append("						        </span>&nbsp;:                                                                                                                                                          \n");
            htmlStr.Append("			</font>                                                                                                                                                                                     \n");
            htmlStr.Append("			</td>                                                                                                                                                                                       \n");
            htmlStr.Append("			<td colspan=4 class=xl13926793 style='border-left: none'>" + dt.Rows[0]["vatamount_display"] + currency + " &nbsp;</td>                                                                                                           \n");
            htmlStr.Append("			<td class=xl7826793>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
            htmlStr.Append("		<tr class=xl6426793 height=26                                                                                                                                                                   \n");
            htmlStr.Append("			style='mso-height-source: userset; height: 20.1pt'>                                                                                                                                         \n");
            htmlStr.Append("			<td height=26 class=xl7326793 style='height: 20.1pt'>&nbsp;</td>                                                                                                                            \n");
            htmlStr.Append("			<td class=xl9426793 style='border-top: none; border-left: none' colspan=11>&nbsp;" + amount_trans + "</td>                                                                                                                 \n");
            htmlStr.Append("			<td class=xl8226793 style='border-top: none'>&nbsp;</td>                                                                                                                                    \n");
            htmlStr.Append("			<td colspan=9 class=xl13626793 style='border-right:1pt solid black'>&nbsp;T&#7893;ng                                                                                                      \n");
            htmlStr.Append("				ti&#7873;n thanh toán (<font class='font826793'>Total payment</font><font                                                                                                               \n");
            htmlStr.Append("				class='font726793'>) :</font>                                                                                                                                                            \n");
            htmlStr.Append("			</td>                                                                                                                                                                                       \n");
            htmlStr.Append("			<td colspan=4 class=xl13926793 style='border-left: none'>" + dt.Rows[0]["totalamount_display"] + currency + " &nbsp;</td>                                                                                                    \n");
            htmlStr.Append("			<td class=xl7826793>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
            htmlStr.Append("		<tr class=xl6426793 height=42                                                                                                                                                                   \n");
            htmlStr.Append("			style='mso-height-source: userset; height: 32.30pt'>                                                                                                                                        \n");
            htmlStr.Append("			<td height=42 class=xl7226793 style='height: 32.30pt'>&nbsp;</td>                                                                                                                           \n");
            htmlStr.Append("			<td colspan=10 class=xl9326793>S&#7889; ti&#7873;n vi&#7871;t                                                                                                                               \n");
            htmlStr.Append("				b&#7857;ng ch&#7919; (<font class='font526793'>Amount in                                                                                                                                \n");
            htmlStr.Append("					words</font><font class='font626793'>) :<span                                                                                                                                       \n");
            htmlStr.Append("					style='mso-spacerun: yes'> </span></font>                                                                                                                                           \n");
            htmlStr.Append("			</td>                                                                                                                                                                                       \n");
            htmlStr.Append("			<td class=xl9326793 style='border-top: none'>&nbsp;</td>                                                                                                                                    \n");
            htmlStr.Append("			<td colspan=14 class=xl14326793 width=372 style='width: 281pt'><span                                                                                                                        \n");
            htmlStr.Append("				style='mso-spacerun: yes'>" + read_prive + "  </span></td>                                                                                                                                     \n");
            htmlStr.Append("			<td class=xl7826793>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
            htmlStr.Append("		<tr height=17 style='mso-height-source: userset; height: 12.9pt'>                                                                                                                               \n");
            htmlStr.Append("			<td height=17 class=xl7526793 style='height: 12.9pt'>&nbsp;</td>                                                                                                                            \n");
            htmlStr.Append("			<td colspan=9 class=xl15326793>Ng&#432;&#7901;i mua hàng (<font                                                                                                                             \n");
            htmlStr.Append("				class='font826793'>Buyer</font><font class='font726793'>)</font></td>                                                                                                                   \n");
            htmlStr.Append("			<td class=xl15426793>&nbsp;</td>                                                                                                                                                            \n");
            htmlStr.Append("			<td class=xl15426793>&nbsp;</td>                                                                                                                                                            \n");
            htmlStr.Append("			<td class=xl15426793>&nbsp;</td>                                                                                                                                                            \n");
            htmlStr.Append("			<td colspan=13 class=xl15326793><span style='mso-spacerun: yes'>                                                                                                                            \n");
            htmlStr.Append("			</span>Ng&#432;&#7901;i bán hàng <font class='font826793'>(Seller</font><font                                                                                                               \n");
            htmlStr.Append("				class='font726793'>)</font></td>                                                                                                                                                        \n");
            htmlStr.Append("			<td class=xl8426793>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
            htmlStr.Append("		<tr height=18 style='height: 15.84pt'>                                                                                                                                                           \n");
            htmlStr.Append("			<td height=18 class=xl7126793 style='height: 15.84pt'>&nbsp;</td>                                                                                                                            \n");
            htmlStr.Append("			<td colspan=9 class=xl9626793>(Ký, ghi rõ h&#7885; tên)</td>                                                                                                                                \n");
            htmlStr.Append("			<td class=xl9526793>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("			<td class=xl9526793>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("			<td class=xl9526793>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("			<td colspan=13 class=xl9626793>(Ký, &#273;óng d&#7845;u, ghi rõ                                                                                                                             \n");
            htmlStr.Append("				h&#7885;, tên)</td>                                                                                                                                                                     \n");
            htmlStr.Append("			<td class=xl7726793>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
            htmlStr.Append("		<tr height=18 style='height: 15.84pt'>                                                                                                                                                           \n");
            htmlStr.Append("			<td height=18 class=xl7126793 style='height: 15.84pt'>&nbsp;</td>                                                                                                                            \n");
            htmlStr.Append("			<td colspan=9 class=xl10326793>(Sign &amp; full name)</td>                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl9526793>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("			<td class=xl9526793>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("			<td class=xl9626793>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("			<td colspan=13 class=xl10326793><span style='mso-spacerun: yes'> </span>(Sign,stamp                                                                                                         \n");
            htmlStr.Append("				&amp; full name)</td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl7726793>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                           \n");

            htmlStr.Append("		<tr height=16 style='mso-height-source: userset; height: 20.0pt'>                                                                                                                               \n");
            htmlStr.Append("			<td height=16 class=xl7126793 style='height: 20.0pt'>&nbsp;</td>                                                                                                                            \n");
            htmlStr.Append("			<td class=xl6526793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl7726793>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
            htmlStr.Append("		<tr height=18 style='mso-height-source: userset; height: 14.1pt'>                                                                                                                               \n");
            htmlStr.Append("			<td height=18 class=xl7126793 style='height: 14.1pt'>&nbsp;</td>                                                                                                                            \n");
            htmlStr.Append("			<td class=xl6526793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl9726793 colspan=4>Signature Valid</td>					                                                                                                                        \n");
          
            if (dt.Rows[0]["sign_yn"].ToString() == "Y")
            {
                htmlStr.Append("			<td align=left valign=top><span                                                                                                                                                             \n");
                htmlStr.Append("				style='mso-ignore: vglayout; position: absolute; z-index: 4; margin-left: 24px; margin-top: 1px; width: 81px; height: 54px'><img                                                        \n");
                htmlStr.Append("					width=81 height=54                                                                                                                                                                  \n");
                htmlStr.Append("					src='D:/webproject/e-invoice-ws/02.Web/EInvoice/img/check_signed.png'                                                                                                                \n");
                htmlStr.Append("					v:shapes='Picture_x0020_8'></span>                                                                                                                                                  \n");
                htmlStr.Append("			<![endif]><span style='mso-ignore: vglayout2'>                                                                                                                                              \n");
                htmlStr.Append("					<table cellpadding=0 cellspacing=0>                                                                                                                                                 \n");
                htmlStr.Append("						<tr>                                                                                                                                                                            \n");
                htmlStr.Append("							<td height=18 class=xl9926793 width=39                                                                                                                                      \n");
                htmlStr.Append("								style='height: 14.1pt; width: 29pt'>&nbsp;</td>                                                                                                                         \n");
                htmlStr.Append("						</tr>                                                                                                                                                                           \n");
                htmlStr.Append("					</table>                                                                                                                                                                            \n");
                htmlStr.Append("			</span></td>                                                                                                                                                                                \n");
            }
            else
            {
                htmlStr.Append("							<td height=18 class=xl9926793 width=39                                                                                                                                      \n");
                htmlStr.Append("								style='height: 14.1pt; width: 29pt'>&nbsp;</td>                                                                                                                         \n");
            }
            
            htmlStr.Append("			                                                                                                                                                                                            \n");
            htmlStr.Append("			<td class=xl9926793>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("			<td class=xl9926793>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("			<td class=xl9926793>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("			<td class=xl9926793>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("			<td class=xl9826793>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("			<td class=xl9826793>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("			<td class=xl9826793>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("			<td class=xl10026793>&nbsp;</td>                                                                                                                                                            \n");
            htmlStr.Append("			<td class=xl7726793>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
            htmlStr.Append("		<tr height=21 style='mso-height-source: userset; height: 15.75pt'>                                                                                                                              \n");
            htmlStr.Append("			<td height=21 class=xl7126793 style='height: 15.75pt'>&nbsp;</td>                                                                                                                           \n");
            htmlStr.Append("			<td class=xl6526793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl6726793></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td colspan=13 class=xl10426793 width=370                                                                                                                                                   \n");
            htmlStr.Append("				style='border-right:1pt solid black; width: 279pt'><font                                                                                                                              \n");
            htmlStr.Append("				class='font926793'>&#272;&#432;&#7907;c ký b&#7903;i:</font><font                                                                                                                       \n");
            htmlStr.Append("				class='font1026793'> </font><font class='font1126793'><span                                                                                                                             \n");
            htmlStr.Append("					style='mso-spacerun: yes'> </span>" + dt.Rows[0]["SignedBy"] + "</font></td>                                                                                                                         \n");
            htmlStr.Append("			<td class=xl7726793>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
            htmlStr.Append("		<tr height=20 style='mso-height-source: userset; height: 18.0pt'>                                                                                                                               \n");
            htmlStr.Append("			<td height=20 class=xl7126793 style='height: 18.0pt'>&nbsp;</td>                                                                                                                            \n");
            htmlStr.Append("			<td colspan=12 class=xl15526793>Mã CQT:                                                                                                                             \n");
            htmlStr.Append("				<font class='font1226793'></font>" + dt.Rows[0]["cqt_mccqt_id"] + "</td>                                                                                                                                      \n");
            htmlStr.Append("			<td colspan=13 class=xl10726793                                                                                                                                                             \n");
            htmlStr.Append("				style='border-right:1pt solid black'>Ngày Ký: <font                                                                                                                                   \n");
            htmlStr.Append("				class='font1326793'>" + dt.Rows[0]["SignedDate"] + "</font></td>                                                                                                                                              \n");
            htmlStr.Append("			<td class=xl7726793>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
            htmlStr.Append("		<tr height=22 style='mso-height-source: userset; height: 20.16pt'>                                                                                                                               \n");
            htmlStr.Append("			<td height=22 class=xl7126793 style='height: 20.16pt'>&nbsp;</td>                                                                                                                            \n");
            htmlStr.Append("			<td colspan=13 class=xl15526793>Tra c&#7913;u t&#7841;i Website:                                                                                                                            \n");
            htmlStr.Append("				<font class='font1226793'>" + dt.Rows[0]["WEBSITE_EI"] + "</font>                                                                                                                   \n");
            htmlStr.Append("			</td>                                                                                                                                                                                       \n");
            htmlStr.Append("			<td class=xl6326793 colspan=12>Mã nhận hóa đơn: " + dt.Rows[0]["matracuu"] + " </td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl7726793>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
            htmlStr.Append("		<tr height=18 style='height: 15.84pt'>                                                                                                                                                           \n");
            htmlStr.Append("			<td colspan=27 height=18 class=xl14426793 width=742                                                                                                                                         \n");
            htmlStr.Append("				style='border-right:1pt solid black; height: 15.84pt; width: 559pt'>(C&#7847;n                                                                                                         \n");
            htmlStr.Append("				ki&#7875;m tra, &#273;&#7889;i chi&#7871;u khi l&#7853;p, giao,                                                                                                                         \n");
            htmlStr.Append("				nh&#7853;n hóa &#273;&#417;n)</td>                                                                                                                                                      \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
            htmlStr.Append("		<tr height=18 style='height: 15.84pt'>                                                                                                                                                           \n");
            htmlStr.Append("			<td colspan=27 height=18 class=xl14726793                                                                                                                                                   \n");
            htmlStr.Append("				style='border-right:1pt solid black; height: 15.84pt'>" + dt.Rows[0]["CONTRACT_INFO_EI"] + "</td>                                                                                                                                                        \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
            htmlStr.Append("	</table>                                                                                                                                                                                            \n");


            htmlStr.Append("</body>                                                                                                                                                                     \n");
            htmlStr.Append("</html>                                                                                                                                                                     \n");

            using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"D:\webproject\e-invoice-ws\02.Web\AttachFileText\" + tei_einvoice_m_pk + ".html"))
            {
                file.WriteLine(htmlStr.ToString()); // "sb" is the StringBuilder
            }


            //ESysLib.WriteLogError("SignXml END:  =======================" + v_countNumberOfPages);
            string result = htmlStr.ToString();
           
            for (int p = 1; p <= number_page; p++ )
            {
                string old_title = "Trang " + p.ToString() + "***";
                string new_title = "";
                if (p==1)
                {
                    new_title = "Trang 1 /" + number_page.ToString();
                    result = result.Replace(old_title, new_title);
                }
                else
                {
                    new_title = "tiep theo trang truoc - Trang " + p.ToString() + " / " + number_page.ToString();
                    result = result.Replace(old_title, new_title);
                }
              

            }

            connection.Close();
            connection.Dispose();
            return result + "|" + dt.Rows[0]["templateCode"].ToString().Replace("/", "") + "_" + dt.Rows[0]["InvoiceSerialNo"].ToString().Replace("/", "") + "_" + dt.Rows[0]["InvoiceNumber"];

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