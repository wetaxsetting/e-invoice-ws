using System;
using System.Collections.Generic;
using System.Data;
//using System.Data.OracleClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Xml;
using System.Xml.Serialization;
using System.Globalization;
using System.Xml.Linq;
using iTextSharp.xmp.impl;
using Oracle.DataAccess.Client;
using Oracle.DataAccess.Types;
namespace EInvoice.Company
{
    public class KyungBang_C
    {
        public static string View(string tei_einvoice_m_pk, string tei_company_pk, string dbName)
        {

            /*string _conString = "Data Source={0};User Id={1};Password={2};Unicode=true";
            // string connString = String.Format( "Data Source={0};Password={1};User ID={2};", m_Host, m_User, m_Password); dbName = "NOBLANDBD",
            string dbUser = "genuwin", dbPwd = "genuwin2";//NOBLANDBD  EINVOICE_252
            _conString = String.Format(_conString, dbName, dbUser, dbPwd);*/
            string _conString = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=123.30.104.243)(PORT=1941))(CONNECT_DATA=(SERVER=dedicated)(SERVICE_NAME=NOBLANDBD)));User ID=genuwin;Password=genuwin2";

            int result = 0; string xml_json = "";

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

            int pos = 10, pos_lv = 20, v_count = 0, count_page = 0, count_page_v = 0, r = 0, x = 0;
            v_count = dt_d.Rows.Count;  //_Invoices.Inv[0].Invoice.Products.Product.Count();
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

            string read_prive;
            if (dt.Rows[0]["CurrencyCodeUSD"].ToString() == "VND")
            {
                read_prive = NumberToTextVN(Decimal.Parse(dt.Rows[0]["TotalAmountInWord"].ToString()));
            }
            else
            {
                read_prive = Num2VNText(dt.Rows[0]["TotalAmountInWord"].ToString().ToString(), "USD");
            }



            //read_prive = NumberToTextVN(Total_Amount_d);
            read_prive = read_prive.Replace(",", "");

            read_prive = read_prive.ToString().Substring(0, 2) + read_prive.ToString().Substring(2, read_prive.Length - 2).ToLower().Replace("mỹ", "Mỹ");

            int end = 0;
            int count = count_page_v + r;
            double height = 130;
            StringBuilder htmlStr = new StringBuilder("");
            string heigh = "", heigh_d = "";

            htmlStr.Append("<html xmlns:v='urn:schemas-microsoft-com:vml'																										\n");
            htmlStr.Append("xmlns:o='urn:schemas-microsoft-com:office:office'                                                                                                  \n");
            htmlStr.Append("xmlns:x='urn:schemas-microsoft-com:office:excel'                                                                                                   \n");
            htmlStr.Append("xmlns='http://www.w3.org/TR/REC-html40'>                                                                                                           \n");
            htmlStr.Append(" 																																						 \n");
            htmlStr.Append("<head>                                                                                                                                             \n");
            htmlStr.Append("<meta charset='UTF-8'>\n");
            htmlStr.Append("<meta http-equiv='Content-Type' content='text/html;charset=UTF-8'>                                                                                      \n");
            htmlStr.Append("<meta name=ProgId content=Excel.Sheet>                                                                                                             \n");
            htmlStr.Append("<meta name=Generator content='Microsoft Excel 15'>                                                                                                 \n");
            htmlStr.Append("<script type='text/javascript' src='${pageContext.request.contextPath}/system/syscommand.js'></script>                                                   \n");
            htmlStr.Append("<title>Report E-Invoice</title>                                                                                                                          \n");
            htmlStr.Append("<!-- Normalize or reset CSS with your favorite library -->                                                                                               \n");
            // htmlStr.Append("<link rel='stylesheet'href='https://cdnjs.cloudflare.com/ajax/libs/normalize/3.0.3/normalize.css'>                                                       \n");
            htmlStr.Append("<!-- Load paper.css for happy printing -->                                                                                                               \n");
            // htmlStr.Append("<link rel='stylesheet' href='https://cdnjs.cloudflare.com/ajax/libs/paper-css/0.2.3/paper.css'>                                                           \n");
            htmlStr.Append("<!-- Set page size here: A5, A4 or A3 -->                                                                                                                \n");
            htmlStr.Append("<!-- Set also 'landscape' if you need -->                                                                                                                \n");
            htmlStr.Append("<style>                                                                                                                                                  \n");
            htmlStr.Append("@page {                                                                                                                                                  \n");
            htmlStr.Append("	size: A4                                                                                                                                             \n");
            htmlStr.Append("}                                                                                                                                                        \n");
            htmlStr.Append("v\\:* {behavior:url(#default#VML);}                                                                                                                 \n");
            htmlStr.Append("o\\:* {behavior:url(#default#VML);}                                                                                                                 \n");
            htmlStr.Append("x\\:* {behavior:url(#default#VML);}                                                                                                                 \n");
            htmlStr.Append(".shape {behavior:url(#default#VML);}                                                                                                               \n");
            htmlStr.Append("</style>                                                                                                                                                 \n");
            htmlStr.Append("<style>                                                                                                                                                  \n");
            htmlStr.Append("body {                                                                                                                                                   \n");
            htmlStr.Append("	color: blue;                                                                                                                                         \n");
            htmlStr.Append("	font-size: 140%;                                                                                                                                     \n");
            htmlStr.Append("	background-image: url('assets/Solution.jpg');                                                                                                        \n");
            htmlStr.Append("	background: white;                                                                                                        \n");
            htmlStr.Append("	width: 100%;                                                                                                       \n");
            htmlStr.Append("	height: 120%;                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                         \n");
            htmlStr.Append("h1 {                                                                                                                                                     \n");
            htmlStr.Append("	color: #00FF00;                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                         \n");
            htmlStr.Append("p {                                                                                                                                                      \n");
            htmlStr.Append("	color: rgb(0, 0, 255)                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                         \n");
            htmlStr.Append("headline1 {                                                                                                                                              \n");
            htmlStr.Append("	background-image: url(assets/Solution.jpg);                                                                                                          \n");
            htmlStr.Append("	background-repeat: no-repeat;                                                                                                                        \n");
            htmlStr.Append("	background-position: left top;                                                                                                                       \n");
            htmlStr.Append("	padding-top: 68px;                                                                                                                                   \n");
            htmlStr.Append("	margin-bottom: 50px;                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                         \n");
            htmlStr.Append("headline2 {                                                                                                                                              \n");
            htmlStr.Append("	background-image: url(images/newsletter_headline2.gif);                                                                                              \n");
            htmlStr.Append("	background-repeat: no-repeat;                                                                                                                        \n");
            htmlStr.Append("	background-position: left top;                                                                                                                       \n");
            htmlStr.Append("	padding-top: 68px;                                                                                                                                   \n");
            htmlStr.Append("}                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                         \n");
            htmlStr.Append("<!--table                                                                                                                                                \n");
            htmlStr.Append("	{mso-displayed-decimal-separator:'\\.';                                                                                                               \n");
            htmlStr.Append("	mso-displayed-thousand-separator:'\\, ';}                                                                                                              \n");
            htmlStr.Append(".font99999                                                                                                                                               \n");
            htmlStr.Append("	{color:windowtext;                                                                                                                                   \n");
            htmlStr.Append("	font-size:22.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	color: red;}                                                                                                                                         \n");
            htmlStr.Append(".font511323                                                                                                                                              \n");
            htmlStr.Append("	{color:windowtext;                                                                                                                                   \n");
            htmlStr.Append("	font-size:18.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                 \n");
            htmlStr.Append(".font611323                                                                                                                                              \n");
            htmlStr.Append("	{color:windowtext;                                                                                                                                   \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                 \n");
            htmlStr.Append(".font711323                                                                                                                                              \n");
            htmlStr.Append("	{color:windowtext;                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                 \n");
            htmlStr.Append(".font811323                                                                                                                                              \n");
            htmlStr.Append("	{color:windowtext;                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                 \n");
            htmlStr.Append(".font911323                                                                                                                                              \n");
            htmlStr.Append("	{color:windowtext;                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                 \n");
            htmlStr.Append(".font1011323                                                                                                                                             \n");
            htmlStr.Append("	{color:windowtext;                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                      \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                 \n");
            htmlStr.Append(".font1111323                                                                                                                                             \n");
            htmlStr.Append("	{color:windowtext;                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                 \n");
            htmlStr.Append(".font1211323                                                                                                                                             \n");
            htmlStr.Append("	{color:windowtext;                                                                                                                                   \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                 \n");
            htmlStr.Append("                                                                                                                                                         \n");
            htmlStr.Append(".font12113233                                                                                                                                            \n");
            htmlStr.Append("	{color:windowtext;                                                                                                                                   \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                 \n");
            htmlStr.Append(".font1311323                                                                                                                                             \n");
            htmlStr.Append("	{color:windowtext;                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                 \n");
            htmlStr.Append(".font1411323                                                                                                                                             \n");
            htmlStr.Append("	{color:windowtext;                                                                                                                                   \n");
            htmlStr.Append("	font-size:16.25pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                 \n");
            htmlStr.Append(".font1511323                                                                                                                                             \n");
            htmlStr.Append("	{color:red;                                                                                                                                          \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                 \n");
            htmlStr.Append(".font1611323                                                                                                                                             \n");
            htmlStr.Append("	{color:red;                                                                                                                                          \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                 \n");
            htmlStr.Append(".font1711323                                                                                                                                             \n");
            htmlStr.Append("	{color:red;                                                                                                                                          \n");
            htmlStr.Append("	font-size:9pt;                                                                                                                                       \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                 \n");
            htmlStr.Append(".font1811323                                                                                                                                             \n");
            htmlStr.Append("	{color:red;                                                                                                                                          \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                 \n");
            htmlStr.Append(".font1911323                                                                                                                                             \n");
            htmlStr.Append("	{color:#0066CC;                                                                                                                                      \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                 \n");
            htmlStr.Append(".font2011323                                                                                                                                             \n");
            htmlStr.Append("	{color:red;                                                                                                                                          \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                 \n");
            htmlStr.Append(".font2111323                                                                                                                                             \n");
            htmlStr.Append("	{color:#003366;                                                                                                                                      \n");
            htmlStr.Append("	font-size:16.25pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:Calibri, sans-serif;                                                                                                                     \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                 \n");
            htmlStr.Append(".font2211323                                                                                                                                             \n");
            htmlStr.Append("	{color:#003366;                                                                                                                                      \n");
            htmlStr.Append("	font-size:18.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:Calibri, sans-serif;                                                                                                                     \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                 \n");
            htmlStr.Append(".font2311323                                                                                                                                             \n");
            htmlStr.Append("	{color:#FF6600;                                                                                                                                      \n");
            htmlStr.Append("	font-size:16.25pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                 \n");
            htmlStr.Append(".font2411323                                                                                                                                             \n");
            htmlStr.Append("	{color:red;                                                                                                                                          \n");
            htmlStr.Append("	font-size:14.0pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:Calibri, sans-serif;                                                                                                                     \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                 \n");
            htmlStr.Append(".xl6511323                                                                                                                                               \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	mso-background-source:auto;                                                                                                                          \n");
            htmlStr.Append("	mso-pattern:auto;                                                                                                                                    \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl6611323                                                                                                                                               \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl6711323                                                                                                                                               \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl6811323                                                                                                                                               \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl6911323                                                                                                                                               \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl7011323                                                                                                                                               \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl7111323                                                                                                                                               \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:none;                                                                                                                                     \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl7211323                                                                                                                                               \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:18.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl7311323                                                                                                                                               \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:none;                                                                                                                                     \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl7411323                                                                                                                                               \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:14.0pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl7511323                                                                                                                                               \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:14.0pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl7611323                                                                                                                                               \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl7711323                                                                                                                                               \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:16.25pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl7811323                                                                                                                                               \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                              \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl7911323                                                                                                                                               \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:18.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl8011323                                                                                                                                               \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl8111323                                                                                                                                               \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:none;                                                                                                                                     \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl8211323                                                                                                                                               \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:16.25pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:right;                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl8311323                                                                                                                                               \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:none;                                                                                                                                     \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl8411323                                                                                                                                               \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:18.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:center;                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl8511323                                                                                                                                               \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:center;                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	background: white;ground:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl8611323                                                                                                                                               \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:center;                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl8711323                                                                                                                                               \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:18.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl8811323                                                                                                                                               \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:18.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                              \n");
            htmlStr.Append("	text-align:right;                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                  \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                       \n");
            htmlStr.Append(".xl8911323                                                                                                                                               \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl9011323                                                                                                                                               \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:none;                                                                                                                                     \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl9111323                                                                                                                                               \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:none;                                                                                                                                     \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl9211323                                                                                                                                               \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:none;                                                                                                                                     \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl9311323                                                                                                                                               \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:18.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:none;                                                                                                                                     \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl9411323                                                                                                                                               \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:none;                                                                                                                                     \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl9511323                                                                                                                                               \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:none;                                                                                                                                     \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl9611323                                                                                                                                               \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:none;                                                                                                                                     \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl9711323                                                                                                                                               \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:none;                                                                                                                                     \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl9811323                                                                                                                                               \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:18.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:none;                                                                                                                                     \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl9911323                                                                                                                                               \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl10011323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:18.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:none;                                                                                                                                     \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl10111323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:#C00000;                                                                                                                                       \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl10211323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:18.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:none;                                                                                                                                     \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl10311323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl10411323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl10511323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:22.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl10611323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:22.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl10711323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:red;                                                                                                                                           \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl10811323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:16.25pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:bottom;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl10911323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:16.25pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:bottom;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl11011323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:16.25pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:bottom;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl11111323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:red;                                                                                                                                           \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:none;                                                                                                                                     \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl11211323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:16.25pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:bottom;                                                                                                                               \n");
            htmlStr.Append("	border-top:none;                                                                                                                                     \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl11311323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:red;                                                                                                                                           \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:none;                                                                                                                                     \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl11411323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:red;                                                                                                                                           \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:none;                                                                                                                                     \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl11511323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:red;                                                                                                                                           \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:none;                                                                                                                                     \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl11611323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:#0070C0;                                                                                                                                       \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl11711323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:none;                                                                                                                                     \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                  \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                       \n");
            htmlStr.Append(".xl11811323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                  \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                       \n");
            htmlStr.Append(".xl11911323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:18.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                  \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                       \n");
            htmlStr.Append(".xl12011323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:red;                                                                                                                                           \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:none;                                                                                                                                     \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                  \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                       \n");
            htmlStr.Append(".xl12111323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:22.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl12211323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:18.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:center;                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl12311323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:left;                                                                                                                                     \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl12411323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:center;                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl12511323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:left;                                                                                                                                     \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl12611323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:none;                                                                                                                                     \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl12711323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:left;                                                                                                                                     \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:none;                                                                                                                                     \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                  \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                       \n");
            htmlStr.Append(".xl12811323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                              \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt dotted  windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:1pt dotted  windowtext;                                                                                                               \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                  \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                       \n");
            htmlStr.Append(".xl12911323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                              \n");
            htmlStr.Append("	text-align:right;                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt dotted  windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:1pt dotted  windowtext;                                                                                                               \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                  \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                       \n");
            htmlStr.Append(".xl13011323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                              \n");
            htmlStr.Append("	text-align:left;                                                                                                                                     \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt dotted  windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl13111323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:18.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                              \n");
            htmlStr.Append("	text-align:right;                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                  \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                       \n");
            htmlStr.Append(".xl13211323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:center;                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt dotted  windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-right:1pt solid ;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:1pt dotted  ;                                                                                                               \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                  \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                       \n");
            htmlStr.Append(".xl13311323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                              \n");
            htmlStr.Append("	text-align:right;                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt dotted  windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:1pt dotted  windowtext;                                                                                                               \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                  \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                       \n");
            htmlStr.Append(".xl13411323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                              \n");
            htmlStr.Append("	text-align:right;                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt dotted  windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:1pt dotted  windowtext;                                                                                                               \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                  \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                       \n");
            htmlStr.Append(".xl13511323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                              \n");
            htmlStr.Append("	text-align:center;                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt dotted  windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:1pt dotted  windowtext;                                                                                                               \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                  \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                       \n");
            htmlStr.Append(".xl13611323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                              \n");
            htmlStr.Append("	text-align:center;                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt dotted  windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:1pt dotted  windowtext;                                                                                                               \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                  \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                       \n");
            htmlStr.Append(".xl13711323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                              \n");
            htmlStr.Append("	text-align:center;                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt dotted  windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:1pt dotted  windowtext;                                                                                                               \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                  \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                       \n");
            htmlStr.Append(".xl13811323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:left;                                                                                                                                     \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:1pt dotted  windowtext;                                                                                                               \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl13911323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:18.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:left;                                                                                                                                     \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt dotted  windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:1pt dotted  windowtext;                                                                                                               \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl14011323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:18.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:left;                                                                                                                                     \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt dotted  windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:1pt dotted  windowtext;                                                                                                               \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl14111323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:center;                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:pt dotted  windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl14211323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                              \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt dotted  windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl14311323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                              \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt dotted  windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl14411323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                              \n");
            htmlStr.Append("	text-align:right;                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt dotted  windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl14511323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                              \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt dotted  windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl14611323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl14711323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:left;                                                                                                                                     \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl14811323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                              \n");
            htmlStr.Append("	text-align:center;                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt dotted  windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:1pt dotted  windowtext;                                                                                                               \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                  \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                       \n");
            htmlStr.Append(".xl14911323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                              \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt dotted  windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl15011323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:red;                                                                                                                                           \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:general;                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:bottom;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl15111323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:0%;                                                                                                                                \n");
            htmlStr.Append("	text-align:left;                                                                                                                                     \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl15211323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:left;                                                                                                                                     \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                  \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                       \n");
            htmlStr.Append(".xl15311323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:left;                                                                                                                                     \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                  \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                       \n");
            htmlStr.Append(".xl15411323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:left;                                                                                                                                     \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                  \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                       \n");
            htmlStr.Append(".xl15511323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:'\\#\\,\\#\\#0';                                                                                                                       \n");
            htmlStr.Append("	text-align:right;                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt dotted  windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:1pt dotted  windowtext;                                                                                                               \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl15611323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:'\\#\\,\\#\\#0';                                                                                                                       \n");
            htmlStr.Append("	text-align:right;                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt dotted  windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:1pt dotted  windowtext;                                                                                                               \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl15711323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:center;                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt dotted  windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:1pt dotted  windowtext;                                                                                                               \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl15811323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:center;                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt dotted  windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:1pt dotted  windowtext;                                                                                                               \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl15911323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                              \n");
            htmlStr.Append("	text-align:right;                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt dotted  windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                  \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                       \n");
            htmlStr.Append(".xl16011323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                              \n");
            htmlStr.Append("	text-align:right;                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt dotted  windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                  \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                       \n");
            htmlStr.Append(".xl16111323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:center;                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt dotted  windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl16211323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:center;                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt dotted  windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl16311323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:red;                                                                                                                                           \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:left;                                                                                                                                     \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:none;                                                                                                                                     \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;                                                                                                                                  \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                       \n");
            htmlStr.Append(".xl16411323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:red;                                                                                                                                           \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:left;                                                                                                                                     \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                  \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                       \n");
            htmlStr.Append(".xl16511323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:red;                                                                                                                                           \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:left;                                                                                                                                     \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:none;                                                                                                                                     \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                  \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                       \n");
            htmlStr.Append(".xl16611323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:center;                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl16711323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:18.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:center;                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl16811323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:18.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:center;                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl16911323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:18.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:left;                                                                                                                                     \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl17011323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:center;                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl17111323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                              \n");
            htmlStr.Append("	text-align:right;                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                  \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                       \n");
            htmlStr.Append(".xl17211323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:18.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:left;                                                                                                                                     \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt dotted  windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl17311323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:center;                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl138113233                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:20.0pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:left;                                                                                                                                     \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:1pt dotted  windowtext;                                                                                                               \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;                                                                                                                                  \n");
            htmlStr.Append("	color: red;}                                                                                                                                         \n");
            htmlStr.Append(".xl17411323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:center;                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl17511323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:center;                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:none;                                                                                                                                     \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl17611323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:18.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:center;                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl17711323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:18.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:center;                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl17811323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:center;                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:none;                                                                                                                                     \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl17911323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:center;                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:none;                                                                                                                                     \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl18011323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:red;                                                                                                                                           \n");
            htmlStr.Append("	font-size:21.25pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:center;                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                  \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl18111323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:18.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:center;                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl18211323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:center;                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:none;                                                                                                                                     \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl18311323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:center;                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:none;                                                                                                                                     \n");
            htmlStr.Append("	border-right:1pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl18411323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:left;                                                                                                                                     \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl18511323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:red;                                                                                                                                           \n");
            htmlStr.Append("	font-size:14.0pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                     \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:center;                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                 \n");
            htmlStr.Append(".xl18611323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                           \n");
            htmlStr.Append("	text-align:left;                                                                                                                                     \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:1pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                 \n");
            htmlStr.Append(".xl18711323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                              \n");
            htmlStr.Append("	text-align:right;                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:1pt dotted  windowtext;                                                                                                               \n");
            htmlStr.Append("	border-left:1pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                  \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                       \n");
            htmlStr.Append(".xl18811323                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                  \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                     \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                   \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                  \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                              \n");
            htmlStr.Append("	text-align:right;                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                               \n");
            htmlStr.Append("	border-top:1pt solid windowtext;                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                   \n");
            htmlStr.Append("	border-bottom:1pt dotted  windowtext;                                                                                                               \n");
            htmlStr.Append("	border-left:none;                                                                                                                                    \n");
            htmlStr.Append("	background:white;                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                              \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                  \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                       \n");
            htmlStr.Append("	.xl6526793 {                                                                                                                                           \n");
            htmlStr.Append("	    padding: 0px;                                                                                                                                      \n");
            htmlStr.Append("	    mso-ignore: padding;                                                                                                                               \n");
            htmlStr.Append("	    color: windowtext;                                                                                                                                 \n");
            htmlStr.Append("	    font-size: 12.0pt;                                                                                                                                 \n");
            htmlStr.Append("	    font-weight: 400;                                                                                                                                  \n");
            htmlStr.Append("	    font-style: normal;                                                                                                                                \n");
            htmlStr.Append("	    text-decoration: none;                                                                                                                             \n");
            htmlStr.Append("	    font-family: 'Times New Roman', serif;                                                                                                             \n");
            htmlStr.Append("	    mso-font-charset: 0;                                                                                                                               \n");
            htmlStr.Append("	    mso-number-format: General;                                                                                                                        \n");
            htmlStr.Append("	    text-align: general;                                                                                                                               \n");
            htmlStr.Append("	    vertical-align: bottom;                                                                                                                            \n");
            htmlStr.Append("	    mso-background-source: auto;                                                                                                                       \n");
            htmlStr.Append("	    mso-pattern: auto;                                                                                                                                 \n");
            htmlStr.Append("	    white-space: nowrap;                                                                                                                               \n");
            htmlStr.Append("	}                                                                                                                                                      \n");
            htmlStr.Append("	.xl6526793 {                                                                                                                                           \n");
            htmlStr.Append("	    padding: 0px;                                                                                                                                      \n");
            htmlStr.Append("	    mso-ignore: padding;                                                                                                                               \n");
            htmlStr.Append("	    color: windowtext;                                                                                                                                 \n");
            htmlStr.Append("	    font-size: 12.0pt;                                                                                                                                 \n");
            htmlStr.Append("	    font-weight: 400;                                                                                                                                  \n");
            htmlStr.Append("	    font-style: normal;                                                                                                                                \n");
            htmlStr.Append("	    text-decoration: none;                                                                                                                             \n");
            htmlStr.Append("	    font-family: 'Times New Roman', serif;                                                                                                             \n");
            htmlStr.Append("	    mso-font-charset: 0;                                                                                                                               \n");
            htmlStr.Append("	    mso-number-format: General;                                                                                                                        \n");
            htmlStr.Append("	    text-align: general;                                                                                                                               \n");
            htmlStr.Append("	    vertical-align: bottom;                                                                                                                            \n");
            htmlStr.Append("	    mso-background-source: auto;                                                                                                                       \n");
            htmlStr.Append("	    mso-pattern: auto;                                                                                                                                 \n");
            htmlStr.Append("	    white-space: nowrap;                                                                                                                               \n");
            htmlStr.Append("	}                                                                                                                                                      \n");
            htmlStr.Append("	.xl6726793 {                                                                                                                                           \n");
            htmlStr.Append("	    padding: 0px;                                                                                                                                      \n");
            htmlStr.Append("	    mso-ignore: padding;                                                                                                                               \n");
            htmlStr.Append("	    color: windowtext;                                                                                                                                 \n");
            htmlStr.Append("	    font-size: 13.0pt;                                                                                                                                 \n");
            htmlStr.Append("	    font-weight: 400;                                                                                                                                  \n");
            htmlStr.Append("	    font-style: normal;                                                                                                                                \n");
            htmlStr.Append("	    text-decoration: none;                                                                                                                             \n");
            htmlStr.Append("	    font-family: 'Times New Roman', serif;                                                                                                             \n");
            htmlStr.Append("	    mso-font-charset: 0;                                                                                                                               \n");
            htmlStr.Append("	    mso-number-format: General;                                                                                                                        \n");
            htmlStr.Append("	    text-align: general;                                                                                                                               \n");
            htmlStr.Append("	    vertical-align: bottom;                                                                                                                            \n");
            htmlStr.Append("	    mso-background-source: auto;                                                                                                                       \n");
            htmlStr.Append("	    mso-pattern: auto;                                                                                                                                 \n");
            htmlStr.Append("	    white-space: nowrap;                                                                                                                               \n");
            htmlStr.Append("	}                                                                                                                                                      \n");
            htmlStr.Append("	.xl7726793 {                                                                                                                                           \n");
            htmlStr.Append("	    padding: 0px;                                                                                                                                      \n");
            htmlStr.Append("	    mso-ignore: padding;                                                                                                                               \n");
            htmlStr.Append("	    color: windowtext;                                                                                                                                 \n");
            htmlStr.Append("	    font-size: 10.0pt;                                                                                                                                 \n");
            htmlStr.Append("	    font-weight: 400;                                                                                                                                  \n");
            htmlStr.Append("	    font-style: normal;                                                                                                                                \n");
            htmlStr.Append("	    text-decoration: none;                                                                                                                             \n");
            htmlStr.Append("	    font-family: 'Times New Roman', serif;                                                                                                             \n");
            htmlStr.Append("	    mso-font-charset: 0;                                                                                                                               \n");
            htmlStr.Append("	    mso-number-format: General;                                                                                                                        \n");
            htmlStr.Append("	    text-align: general;                                                                                                                               \n");
            htmlStr.Append("	    vertical-align: bottom;                                                                                                                            \n");
            htmlStr.Append("	    border-top: none;                                                                                                                                  \n");
            htmlStr.Append("	    border-right: .5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("	    border-bottom: none;                                                                                                                               \n");
            htmlStr.Append("	    border-left: none;                                                                                                                                 \n");
            htmlStr.Append("	    mso-background-source: auto;                                                                                                                       \n");
            htmlStr.Append("	    mso-pattern: auto;                                                                                                                                 \n");
            htmlStr.Append("	    white-space: nowrap;                                                                                                                               \n");
            htmlStr.Append("	}                                                                                                                                                      \n");
            htmlStr.Append("	.xl14726793 {                                                                                                                                          \n");
            htmlStr.Append("	    padding: 0px;                                                                                                                                      \n");
            htmlStr.Append("	    mso-ignore: padding;                                                                                                                               \n");
            htmlStr.Append("	    color: windowtext;                                                                                                                                 \n");
            htmlStr.Append("	    font-size: 10.0pt;                                                                                                                                 \n");
            htmlStr.Append("	    font-weight: 400;                                                                                                                                  \n");
            htmlStr.Append("	    font-style: italic;                                                                                                                                \n");
            htmlStr.Append("	    text-decoration: none;                                                                                                                             \n");
            htmlStr.Append("	    font-family: 'Times New Roman', serif;                                                                                                             \n");
            htmlStr.Append("	    mso-font-charset: 0;                                                                                                                               \n");
            htmlStr.Append("	    mso-number-format: General;                                                                                                                        \n");
            htmlStr.Append("	    text-align: center;                                                                                                                                \n");
            htmlStr.Append("	    vertical-align: bottom;                                                                                                                            \n");
            htmlStr.Append("	    border-top: none;                                                                                                                                  \n");
            htmlStr.Append("	    border-right: none;                                                                                                                                \n");
            htmlStr.Append("	    border-bottom: .5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("	    border-left: .5pt solid windowtext;                                                                                                                \n");
            htmlStr.Append("	    mso-background-source: auto;                                                                                                                       \n");
            htmlStr.Append("	    mso-pattern: auto;                                                                                                                                 \n");
            htmlStr.Append("	    white-space: nowrap;                                                                                                                               \n");
            htmlStr.Append("	}                                                                                                                                                      \n");
            htmlStr.Append("-->                                                                                                                                                      \n");
            htmlStr.Append("</style>                                                                                                                                                 \n");
            htmlStr.Append("                                                                                                                                                         \n");
            htmlStr.Append("</head>                                                                                                                                                  \n");
            htmlStr.Append("<body class='A4'>                                                                                                                                        \n");
            // htmlStr.Append("	<body>                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                         \n");


            htmlStr.Append("<table border=0 cellpadding=0 cellspacing=0 width=737 align=center margin-top:10px class=xl6611323                                                                                    \n");
            htmlStr.Append(" style='border-collapse:collapse;table-layout:fixed;width:660pt'>                                                                                        \n");
            htmlStr.Append(" <col class=xl6611323 width=6 style='mso-width-source:userset;mso-width-alt:                                                                             \n");
            htmlStr.Append(" 199;width:5pt'>                                                                                                                                         \n");
            htmlStr.Append(" <col class=xl6611323 width=33 style='mso-width-source:userset;mso-width-alt:                                                                            \n");
            htmlStr.Append(" 1166;width:31‬pt'>                                                                                                                                       \n");
            htmlStr.Append(" <col class=xl6611323 width=69 style='mso-width-source:userset;mso-width-alt:                                                                            \n");
            htmlStr.Append(" 2446;width:62pt'>                                                                                                                                       \n");
            htmlStr.Append(" <col class=xl6611323 width=54 style='mso-width-source:userset;mso-width-alt:                                                                            \n");
            htmlStr.Append(" 1934;width:31pt'>                                                                                                                                       \n");
            htmlStr.Append(" <col class=xl6611323 width=40 style='mso-width-source:userset;mso-width-alt:                                                                            \n");
            htmlStr.Append(" 1422;width:50pt'>                                                                                                                                       \n");
            htmlStr.Append(" <col class=xl6611323 width=98 style='mso-width-source:userset;mso-width-alt:                                                                            \n");
            htmlStr.Append(" 3470;width:70‬pt'>                                                                                                                                       \n");
            htmlStr.Append(" <col class=xl6611323 width=27 style='mso-width-source:userset;mso-width-alt:                                                                            \n");
            htmlStr.Append(" 967;width:18pt'>                                                                                                                                        \n");
            htmlStr.Append(" <col class=xl6611323 width=78 style='mso-width-source:userset;mso-width-alt:                                                                            \n");
            htmlStr.Append(" 2759;width:70pt'>                                                                                                                                       \n");
            htmlStr.Append(" <col class=xl6611323 width=56 style='mso-width-source:userset;mso-width-alt:                                                                            \n");
            htmlStr.Append(" 1991;width:52pt'>                                                                                                                                       \n");
            htmlStr.Append(" <col class=xl6611323 width=30 style='mso-width-source:userset;mso-width-alt:                                                                            \n");
            htmlStr.Append(" 1080;width:28pt'>                                                                                                                                       \n");
            htmlStr.Append(" <col class=xl6611323 width=6 style='mso-width-source:userset;mso-width-alt:                                                                             \n");
            htmlStr.Append(" 199;width:5pt'>                                                                                                                                         \n");
            htmlStr.Append(" <col class=xl6611323 width=49 style='mso-width-source:userset;mso-width-alt:                                                                            \n");
            htmlStr.Append(" 1735;width:46pt'>                                                                                                                                       \n");
            htmlStr.Append(" <col class=xl6611323 width=47 style='mso-width-source:userset;mso-width-alt:                                                                            \n");
            htmlStr.Append(" 1678;width:43pt'>                                                                                                                                       \n");
            htmlStr.Append(" <col class=xl6611323 width=6 style='mso-width-source:userset;mso-width-alt:                                                                             \n");
            htmlStr.Append(" 199;width:5pt'>                                                                                                                                         \n");
            htmlStr.Append(" <col class=xl6611323 width=54 style='mso-width-source:userset;mso-width-alt:                                                                            \n");
            htmlStr.Append(" 1934;width:51pt'>                                                                                                                                       \n");
            htmlStr.Append(" <col class=xl6611323 width=78 style='mso-width-source:userset;mso-width-alt:                                                                            \n");
            htmlStr.Append(" 2759;width:75pt'>                                                                                                                                       \n");
            htmlStr.Append(" <col class=xl6611323 width=6 style='mso-width-source:userset;mso-width-alt:                                                                             \n");
            htmlStr.Append(" 199;width:5pt'>                                                                                                                                       \n");
            htmlStr.Append(" <tr height=33 style='mso-height-source:userset;height:31.25‬pt'>                                                                                            \n");
            htmlStr.Append("  <td height=33 class=xl12111323 width=6 style='height:31.25‬pt;width:5pt'>&nbsp;</td>                                                                       \n");
            htmlStr.Append("  <td class=xl6811323 width=33 style='width:31.25‬pt'>&nbsp;</td>                                                                                            \n");
            htmlStr.Append("  <td width=69 style='width:52pt;border-top:1pt solid black' align=left valign=top><![if !vml]><span style='mso-ignore:vglayout;                        \n");
            htmlStr.Append("  position:absolute;z-index:2;margin-left:2px;margin-top:2px;width:123px;                                                                                \n");
            htmlStr.Append("  height:103px'><img width=134 height=118                                                                                                                 \n");
            htmlStr.Append("  src='D:\\webproject\\e-invoice-ws\\02.Web\\EInvoice\\img\\KyungBang_01.png'                                                                                \n");
            htmlStr.Append("  v:shapes='Picture_x0020_1'></span><![endif]><span style='mso-ignore:vglayout2'>                                                                        \n");
            htmlStr.Append("  <table cellpadding=0 cellspacing=0>                                                                                                                    \n");
            htmlStr.Append("   <tr>                                                                                                                                                  \n");
            htmlStr.Append("    <td height=33  width=69 style='height:25.05pt;width:52pt;border-top:none'>&nbsp;</td>                                                                \n");
            htmlStr.Append("   </tr>                                                                                                                                                 \n");
            htmlStr.Append("  </table>                                                                                                                                               \n");
            htmlStr.Append("  </span></td>                                                                                                                                           \n");
            htmlStr.Append("  <td class=xl10511323 width=54 style='width:51.25pt'>&nbsp;</td>                                                                                           \n");
            htmlStr.Append("  <td colspan=8 class=xl18011323 width=384 style='width:287pt'>HÓA                                                                                       \n");
            htmlStr.Append("  ĐƠN GIÁ TRỊ GIA TĂNG</td>                                                                                                                              \n");
            htmlStr.Append("  <td class=xl10511323 width=47 style='width:43.75pt'>&nbsp;</td>                                                                                           \n");
            htmlStr.Append("  <td class=xl10511323 width=6 style='width:5pt'>&nbsp;</td>                                                                                             \n");
            htmlStr.Append("  <td class=xl10511323 width=54 style='width:51.25pt'>&nbsp;</td>                                                                                           \n");
            htmlStr.Append("  <td class=xl10511323 width=78 style='width:72.5‬pt'>&nbsp;</td>                                                                                           \n");
            htmlStr.Append("  <td class=xl10611323 width=6 style='width:5pt'>&nbsp;</td>                                                                                             \n");
            htmlStr.Append(" </tr>                                                                                                                                                   \n");
            htmlStr.Append(" <tr height=24 style='mso-height-source:userset;height:22.5pt'>                                                                                          \n");
            htmlStr.Append("  <td height=24 class=xl12611323 style='height:22.5pt'>&nbsp;</td>                                                                                       \n");
            htmlStr.Append("  <td class=xl6611323>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append("  <td class=xl6611323>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append("  <td class=xl6611323>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append("  <td colspan=8 class=xl18511323>(VAT INVOICE)</td>                                                                                                      \n");
            htmlStr.Append("  <td class=xl7011323 colspan=5 style='border-right:1pt solid black'>M&#7851;u                                                                          \n");
            htmlStr.Append("  s&#7889; (<font class='font1211323'>Form</font><font class='font911323'>): </font><font                                                                \n");
            htmlStr.Append("  class='font811323'>" + dt.Rows[0]["templateCode"] + "</font></td>                                                                                                         \n");
            htmlStr.Append(" </tr>                                                                                                                                                   \n");
            htmlStr.Append(" <tr height=21 style='mso-height-source:userset;height:20.06pt'>                                                                                         \n");
            htmlStr.Append("  <td height=21 class=xl12611323 style='height:20.06pt'>&nbsp;</td>                                                                                      \n");
            htmlStr.Append("  <td class=xl6611323>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append("  <td class=xl6611323>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append("  <td class=xl6611323>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append("  <td colspan=8 class=xl8511323><span                                                                                                                    \n");
            htmlStr.Append("  style='mso-spacerun:yes'> </span>(HÓA &#272;&#416;N CHUY&#7874;N &#272;&#7892;I T&#7914; HÓA &#272;&#416;N &#272;I&#7878;N T&#7916;)</td>                                                                                                                 \n");
            htmlStr.Append("  <td class=xl7011323 colspan=4>Ký hi&#7879;u (<font class='font1211323'>Serial</font><font                                                              \n");
            htmlStr.Append("  class='font911323'>):</font><font class='font811323'> " + dt.Rows[0]["InvoiceSerialNo"] + "</font></td>                                                                    \n");
            htmlStr.Append("  <td class=xl7111323>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append(" </tr>                                                                                                                                                   \n");
            htmlStr.Append(" <tr height=24 style='mso-height-source:userset;height:22.5pt'>                                                                                          \n");
            htmlStr.Append("  <td height=24 class=xl12611323 style='height:22.5pt'>&nbsp;</td>                                                                                       \n");
            htmlStr.Append("  <td class=xl6611323>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append("  <td class=xl6611323>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append("  <td class=xl6611323>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append("  <td colspan=8 class=xl18111323 width=384 style='width:287pt'>Ngày <font                                                                                \n");
            htmlStr.Append("  class='font1211323'>(Date)</font><font class='font1111323'> </font><font                                                                               \n");
            htmlStr.Append("  class='font711323'> " + dt.Rows[0]["invoiceissueddate_dd"] + " tháng </font><font class='font1211323'>(month)</font><font                                                          \n");
            htmlStr.Append("  class='font711323'> " + dt.Rows[0]["invoiceissueddate_mm"] + " n&#259;m </font><font class='font1211323'>(year)</font><font                                                        \n");
            htmlStr.Append("  class='font711323'> " + dt.Rows[0]["invoiceissueddate_yyyy"] + "</font></td>                                                                                                       \n");
            htmlStr.Append("  <td class=xl7011323 colspan=4>S&#7889; (<font class='font1211323'>No</font><font                                                                       \n");
            htmlStr.Append("  class='font911323'>.):</font><font class='font811323'><span                                                                                            \n");
            htmlStr.Append("  style='mso-spacerun:yes'>      </span></font><font class='font2011323'><span                                                                           \n");
            htmlStr.Append("  style='mso-spacerun:yes'> </span></font><font class='font2411323'>" + dt.Rows[0]["InvoiceNumber"] + "</font></td>                                                       \n");
            htmlStr.Append("  <td class=xl7311323 width=6 style='width:5pt'>&nbsp;</td>                                                                                              \n");
            htmlStr.Append(" </tr>                                                                                                                                                   \n");
            htmlStr.Append(" <tr height=24 style='mso-height-source:userset;height:21.25pt'>                                                                                          \n");
            htmlStr.Append("  <td height=24 class=xl12611323 style='height:21.25pt;border-top:1pt solid black'>&nbsp;</td>                                                           \n");
            htmlStr.Append("  <td class=xl7611323 colspan=15 style='border-top:1pt solid black'>&#272;&#417;n v&#7883; bán hàng (<font                                              \n");
            htmlStr.Append("  class='font1211323'>Company name</font><font class='font711323'>): </font><font                                                                        \n");
            htmlStr.Append("  class='font2311323'>" + dt.Rows[0]["Seller_Name"] + "</font></td>                                                                                                     \n");
            htmlStr.Append("  <td class=xl7111323 style='border-top:1pt solid black'>&nbsp;</td>                                                                                    \n");
            htmlStr.Append(" </tr>                                                                                                                                                   \n");
            htmlStr.Append(" <tr height=24 style='mso-height-source:userset;height:21.25pt'>                                                                                          \n");
            htmlStr.Append("  <td height=24 class=xl12611323 style='height:21.25pt'>&nbsp;</td>                                                                                       \n");
            htmlStr.Append("  <td class=xl7611323 colspan=5>Mã s&#7889; thu&#7871; <font class='font1211323'>(Tax code)</font>:<font                                                 \n");
            htmlStr.Append("  class='font1411323'> </font><font class='font2111323'>" + dt.Rows[0]["Seller_TaxCode"] + "</font></td>                                                              \n");
            htmlStr.Append("  <td class=xl7711323>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append("  <td class=xl7611323>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append("  <td class=xl7611323>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append("  <td class=xl7611323>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append("  <td class=xl6611323>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append("  <td class=xl6611323>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append("  <td class=xl6611323>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append("  <td class=xl6611323>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append("  <td class=xl6611323>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append("  <td class=xl6611323>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append("  <td class=xl7111323>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append(" </tr>                                                                                                                                                   \n");
            htmlStr.Append(" <tr height=24 style='mso-height-source:userset;height:21.25pt'>                                                                                          \n");
            htmlStr.Append("  <td height=24 class=xl12611323 style='height:21.25pt'>&nbsp;</td>                                                                                       \n");
            htmlStr.Append("  <td class=xl7611323 colspan=3>&#272;&#7883;a ch&#7881; (<font                                                                                          \n");
            htmlStr.Append("  class='font1211323'>Address</font><font class='font711323'>): </font></td>                                                                             \n");
            htmlStr.Append("  <td class=xl7611323 colspan=12><font                                                                                                                   \n");
            htmlStr.Append("  class='font1011323'>" + dt.Rows[0]["Seller_Address"] + "</font></td>                                                                                                \n");
            htmlStr.Append("  <td class=xl7111323>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append(" </tr>                                                                                                                                                   \n");
            htmlStr.Append(" <tr height=24 style='mso-height-source:userset;height:21.25pt'>                                                                                          \n");
            htmlStr.Append("  <td height=24 class=xl12611323 style='height:21.25pt'>&nbsp;</td>                                                                                       \n");
            htmlStr.Append("  <td class=xl7611323 colspan=15>&#272;i&#7879;n tho&#7841;i <font class='font1211323'>(Tel)</font>: " + dt.Rows[0]["TEL92"] + " - Fax: " + dt.Rows[0]["FAX93"] + "</td> \n");
            htmlStr.Append("                                                                                                                                                         \n");
            htmlStr.Append("  <td class=xl7111323>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append(" </tr>                                                                                                                                                   \n");
            htmlStr.Append(" <tr height=24 style='mso-height-source:userset;height:21.25pt'>                                                                                          \n");
            htmlStr.Append("  <td height=24 class=xl12611323 style='height:21.25pt'>&nbsp;</td>                                                                                       \n");
            htmlStr.Append("  <td class=xl7611323 colspan=15>S&#7889; tài kho&#7843;n (<font                                                                                         \n");
            htmlStr.Append("  class='font1211323'>Acc. code</font><font class='font711323'>):                                                                                        \n");
            htmlStr.Append("  </font>" + dt.Rows[0]["BANK_ACCOUNT188"] + " " + dt.Rows[0]["BANK_NM190"] + "</td>                                                                                                                       \n");
            htmlStr.Append("  <td class=xl7111323>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append(" </tr>                                                                                                                                                   \n");
            htmlStr.Append(" <tr height=24 style='mso-height-source:userset;height:22.5pt'>                                                                                          \n");
            htmlStr.Append("  <td height=24 class=xl12611323 style='height:22.5pt'>&nbsp;</td>                                                                                       \n");
            htmlStr.Append("  <td class=xl7611323 colspan=15>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font             \n");
            htmlStr.Append("  class='font1211323'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font><font class='font711323'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; \n");
            htmlStr.Append("  </font>" + dt.Rows[0]["BANK_ACCOUNT289"] + " " + dt.Rows[0]["BANK_NM291"] + "</td>                                                                                                                     \n");
            htmlStr.Append("  <td class=xl7111323>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append(" </tr>                                                                                                                                                   \n");
            htmlStr.Append(" <tr height=24 style='mso-height-source:userset;height:21.25pt'>                                                                                          \n");
            htmlStr.Append("  <td height=24 class=xl12611323 style='height:21.25pt;border-top:1pt solid black'>&nbsp;</td>                                                           \n");
            htmlStr.Append("  <td class=xl7611323 colspan=5 style='height:21.25pt;border-top:1pt solid black'>H&#7885; tên ng&#432;&#7901;i mua hàng (<font                          \n");
            htmlStr.Append("  class='font1211323'>Customer's name</font><font class='font711323'>):</font></td>                                                                      \n");
            htmlStr.Append("  <td colspan=11 class=xl15411323 style='border-right:1pt solid black;border-top:1pt solid black'>&nbsp;" + dt.Rows[0]["Buyer"] + "</td>                        \n");
            htmlStr.Append(" </tr>                                                                                                                                                   \n");
            htmlStr.Append(" <tr height=24 style='mso-height-source:userset;height:21.25pt'>                                                                                          \n");
            htmlStr.Append("  <td height=24 class=xl12611323 style='height:21.25pt'>&nbsp;</td>                                                                                       \n");
            htmlStr.Append("  <td class=xl7611323 colspan=4>Tên &#273;&#417;n v&#7883; (<font                                                                                        \n");
            htmlStr.Append("  class='font1211323'>Company's name</font><font class='font711323'>):</font></td>                                                                       \n");
            htmlStr.Append("  <td colspan=11 class=xl18411323>" + dt.Rows[0]["BuyerLegalName"] + "</td>                                                                                                \n");
            htmlStr.Append("  <td class=xl12711323>&nbsp;</td>                                                                                                                       \n");
            htmlStr.Append(" </tr>                                                                                                                                                   \n");
            htmlStr.Append(" <tr height=24 style='mso-height-source:userset;height:21.25pt'>                                                                                          \n");
            htmlStr.Append("  <td height=24 class=xl12611323 style='height:21.25pt'>&nbsp;</td>                                                                                       \n");
            htmlStr.Append("  <td class=xl7611323 colspan=15>Mã s&#7889; thu&#7871; <font class='font1211323'>(Tax code)</font>: <font                                               \n");
            htmlStr.Append("  class='font2211323'>" + dt.Rows[0]["BuyerTaxCode"] + " </font></td>                                                                                                      \n");
            htmlStr.Append("  <td class=xl12711323>&nbsp;</td>                                                                                                                       \n");
            htmlStr.Append(" </tr>                                                                                                                                                   \n");
            htmlStr.Append(" <tr height=24 style='mso-height-source:userset;height:21.25pt'>                                                                                          \n");
            htmlStr.Append("  <td height=24 class=xl12611323 style='height:21.25pt'>&nbsp;</td>                                                                                       \n");
            htmlStr.Append("  <td class=xl7611323 colspan=3>&#272;&#7883;a ch&#7881; (<font                                                                                          \n");
            htmlStr.Append("  class='font1211323'>Address</font><font class='font711323'>):                                                                                          \n");
            htmlStr.Append("  </font></td>                                                                                                                                           \n");
            htmlStr.Append("  <td class=xl7611323 colspan=12 style='white-space:normal;'><font class='font711323'>" + dt.Rows[0]["BuyerAddress"] + "                                                                                \n");
            htmlStr.Append("  </font></td>                                                                                                                                           \n");
            htmlStr.Append("  <td class=xl8111323>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append(" </tr>                                                                                                                                                   \n");
            htmlStr.Append(" <tr height=26 style='mso-height-source:userset;height:21.375‬pt'>                                                                                          \n");
            htmlStr.Append("  <td height=26 class=xl12611323 style='height:21.375‬pt'>&nbsp;</td>                                                                                       \n");
            htmlStr.Append("  <td class=xl7611323 colspan=6>Hình th&#7913;c thanh toán (<font                                                                                        \n");
            htmlStr.Append("  class='font1211323'>Mod of payment</font><font class='font711323'>): </font><font                                                                      \n");
            htmlStr.Append("  class='font511323'> " + dt.Rows[0]["PaymentMethodCK"] + " </font></td>                                                                                                \n");
            htmlStr.Append("  <td class=xl7611323>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append("  <td class=xl6611323 colspan=5><font class='font711323'>S&#7889; tài                                                                                    \n");
            htmlStr.Append("  kho&#7843;n (</font><font class='font1111323'>Account code</font><font                                                                                 \n");
            htmlStr.Append("  class='font711323'>):</font></td>                                                                                                                      \n");
            htmlStr.Append("  <td class=xl6611323>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append("  <td class=xl6611323>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append("  <td class=xl6611323>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append("  <td class=xl7111323>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append(" </tr>                                                                                                                                                   \n");
            htmlStr.Append(" <tr class=xl8411323 height=21 style='height:19.5pt'>                                                                                                    \n");
            htmlStr.Append("  <td colspan=2 height=21 class=xl17611323 style='height:19.5pt'>STT</td>                                                                                \n");
            htmlStr.Append("  <td colspan=5 class=xl17611323 style='border-right:1pt solid black'>Tên hàng                                                                          \n");
            htmlStr.Append("  hóa, d&#7883;ch v&#7909;</td>                                                                                                                          \n");
            htmlStr.Append("  <td class=xl12211323>&#272;&#417;n v&#7883; tính</td>                                                                                                  \n");
            htmlStr.Append("  <td colspan=3 class=xl17611323 style='border-right:1pt solid black'>S&#7889;                                                                          \n");
            htmlStr.Append("  l&#432;&#7907;ng</td>                                                                                                                                  \n");
            htmlStr.Append("  <td colspan=3 class=xl12211323>&#272;&#417;n giá</td>                                                                                                  \n");
            htmlStr.Append("  <td colspan=3 class=xl17611323 style='border-right:1pt solid black'>Thành                                                                             \n");
            htmlStr.Append("  ti&#7873;n</td>                                                                                                                                        \n");
            htmlStr.Append(" </tr>                                                                                                                                                   \n");
            htmlStr.Append(" <tr class=xl8511323 height=18 style='height:16.5pt'>                                                                                                    \n");
            htmlStr.Append("  <td colspan=2 height=18 class=xl17811323 style='height:16.5pt'>No.</td>                                                                                \n");
            htmlStr.Append("  <td colspan=5 class=xl18211323 style='border-right:1pt solid black'>Description                                                                       \n");
            htmlStr.Append("  of goods</td>                                                                                                                                          \n");
            htmlStr.Append("  <td class=xl8511323>Unit</td>                                                                                                                          \n");
            htmlStr.Append("  <td colspan=3 class=xl17811323 style='border-right:1pt solid black'>Quantity</td>                                                                     \n");
            htmlStr.Append("  <td colspan=3 class=xl8511323>Unit price</td>                                                                                                          \n");
            htmlStr.Append("  <td colspan=3 class=xl17811323 style='border-right:1pt solid black'>Amount</td>                                                                       \n");
            htmlStr.Append(" </tr>                                                                                                                                                   \n");
            htmlStr.Append(" <tr class=xl8611323 height=20 style='mso-height-source:userset;height:18.75pt'>                                                                          \n");
            htmlStr.Append("  <td colspan=2 height=20 class=xl17311323 style='height:18.75pt'>1</td>                                                                                  \n");
            htmlStr.Append("  <td colspan=5 class=xl17311323 style='border-right:1pt solid black'>2</td>                                                                            \n");
            htmlStr.Append("  <td class=xl12411323>3</td>                                                                                                                            \n");
            htmlStr.Append("  <td colspan=3 class=xl17311323 style='border-right:1pt solid black'>4</td>                                                                            \n");
            htmlStr.Append("  <td colspan=3 class=xl12411323>5</td>                                                                                                                  \n");
            htmlStr.Append("  <td colspan=3 class=xl17311323 style='border-right:1pt solid black'>6 = 4 x                                                                           \n");
            htmlStr.Append("  5</td>                                                                                                                                                 \n");
            htmlStr.Append("  </tr>                                                                                                                                                  \n");

            for (int dtR = 0; dtR < v_count; dtR++)
            {
                if (dt_d.Rows[dtR][0].ToString().Length > 28)
                {
                    heigh = "21.0pt";

                }
                else
                {
                    heigh = "25.0pt";
                }
                if (dtR == 0)
                {


                    htmlStr.Append("  					<tr class=xl8411323 height=25 style='mso-height-source:userset;height:" + heigh + "'>                                                        \n");
                    htmlStr.Append("					  <td colspan=2 height=25 class=xl15711323 width=39 style='border-right:1pt solid black; border-bottom:1pt dotted  windowtext;                                              \n");
                    htmlStr.Append("					  height:" + heigh + ";width:29pt'>" + dt_d.Rows[dtR][7] + "</td>                                                                                        \n");
                    htmlStr.Append("					  <td colspan=5 height=25 width=288 style='border-right:1pt solid black;                                                            \n");
                    htmlStr.Append("							  height:25.0pt;width:216pt' align=left valign=top><span style='mso-ignore:vglayout;                                          \n");
                    htmlStr.Append("							  position:absolute;z-index:3;margin-left:210px;margin-top:25px;width:290px;                                                 \n");
                    htmlStr.Append("							  height:245px'><img width=362 height=306                                                                                    \n");
                    htmlStr.Append("							  src='D:\\webproject\\e-invoice-ws\\02.Web\\EInvoice\\img\\KyungBang_05.png' style='background: none;'                                                    \n");
                    htmlStr.Append("							  v:shapes='Picture_x0020_2'/></span><![endif]><span style='mso-ignore:vglayout2'>                                            \n");
                    htmlStr.Append("							  <table cellpadding=0 cellspacing=0 height=100%>                                                                                        \n");
                    htmlStr.Append("							   <tr>                                                                                                                      \n");
                    htmlStr.Append("							    <td colspan=5 height=25 class=xl13811323 width=288 style='height:18pt;border-left:none;border-bottom:none;width:216pt'>&nbsp;" + dt_d.Rows[dtR][0] + " </td > \n");
                    htmlStr.Append("							         </tr>                                                                                                                        \n");
                    htmlStr.Append("							  </table>                                                                                                                   \n");
                    htmlStr.Append("							  </span></td>                                                                                                               \n");
                    htmlStr.Append("					  <td class=xl13211323 style='border-left:none;border-bottom:1pt dotted  windowtext;'>&nbsp;" + dt_d.Rows[dtR][1] + "</td>                                                             \n");
                    htmlStr.Append("					  <td colspan=2 class=xl18711323 style='border-left:none;border-bottom:1pt dotted  windowtext;'>" + dt_d.Rows[dtR][2] + "</td>                                                         \n");
                    htmlStr.Append("					  <td class=xl12911323>&nbsp;</td>                                                                                                   \n");
                    htmlStr.Append("					  <td colspan=2 class=xl13311323 style='border-left:none'>" + dt_d.Rows[dtR][3] + "</td>                                                         \n");
                    htmlStr.Append("					  <td class=xl12811323>&nbsp;</td>                                                                                                   \n");
                    htmlStr.Append("					  <td colspan=2 class=xl15511323 width=132 style='border-left:none;width:99pt'>" + dt_d.Rows[dtR][4] + "</td>                                    \n");
                    htmlStr.Append("					  <td class=xl12811323>&nbsp;</td>                                                                                                   \n");
                    htmlStr.Append("					</tr>                                                                                                                                \n");
                    htmlStr.Append("  						                                                                                                                                 \n");
                }
                else if (dtR == 10)
                {

                    htmlStr.Append("	  							<tr class=xl7611323 height=25 style='mso-height-source:userset;height:" + heigh + "'>                                            \n");
                    htmlStr.Append("								  <td colspan=2 height=25 class=xl16111323 width=39 style='border-right:1pt solid black;                                \n");
                    htmlStr.Append("								  height:" + heigh + ";width:29pt'>&nbsp;" + dt_d.Rows[dtR][7] + "</td>                                                                      \n");
                    htmlStr.Append("								  <td class=xl14111323 width=69 style='border-top:none;width:52pt;' colspan=2>&nbsp;" + dt_d.Rows[dtR][0] + "</td>                    \n");
                    htmlStr.Append("								  <td colspan=3 class=xl17211323 width=165 style='width:128.75pt'>&nbsp;</td>                                               \n");
                    htmlStr.Append("								  <td class=xl14211323 width=78 style='border-top:none;width:72.5‬pt'>&nbsp;" + dt_d.Rows[dtR][1] + "</td>                              \n");
                    htmlStr.Append("								  <td class=xl14311323 width=56 style='border-top:none;border-left:none; width:52.5pt' colspan=2>&nbsp;" + dt_d.Rows[dtR][2] + "</td>  \n");
                    htmlStr.Append("								  <td class=xl14411323 width=6 style='border-top:none;width:5pt'>&nbsp;</td>                                             \n");
                    htmlStr.Append("								  <td colspan=2 class=xl15911323 style='border-left:none'>&nbsp" + dt_d.Rows[dtR][3] + "</td>                                       \n");
                    htmlStr.Append("								  <td class=xl14511323 width=6 style='border-top:none;width:5pt'>&nbsp;</td>                                             \n");
                    htmlStr.Append("								  <td colspan=2 class=xl15911323 style='border-left:none'>&nbsp;" + dt_d.Rows[dtR][4] + "</td>                                       \n");
                    htmlStr.Append("								  <td class=xl13011323 width=6 style='border-top:none;width:5pt'>&nbsp;</td>                                             \n");
                    htmlStr.Append("								 </tr>                                                                                                                   \n");
                    htmlStr.Append("  						                                                                                                                                 \n");

                }
                else
                {
                    htmlStr.Append("  						    <tr class=xl8411323 height=25 style='mso-height-source:userset;height:" + heigh + "'>                                            \n");
                    htmlStr.Append("							  <td colspan=2 height=25 class=xl15711323 width=39 style='border-right:1pt solid black;                                    \n");
                    htmlStr.Append("							  height:" + heigh + ";width:29pt'>&nbsp;" + dt_d.Rows[dtR][7] + "</td>                                                                          \n");
                    htmlStr.Append("							  <td colspan=5 class=xl13811323 width=288 style='border-right:1pt solid black; border-top:1pt dotted  windowtext;                                            \n");
                    htmlStr.Append("							  border-left:none;width:216pt'>&nbsp;" + dt_d.Rows[dtR][0] + "</td>                                                                     \n");
                    htmlStr.Append("							  <td class=xl13211323 style='border-top:none;border-left:none'>&nbsp;" + dt_d.Rows[dtR][1] + "</td>                                     \n");
                    htmlStr.Append("							  <td class=xl13311323 style='border-top:none;border-left:none' colspan=2>&nbsp;" + dt_d.Rows[dtR][2] + "</td>                           \n");
                    htmlStr.Append("							  <td class=xl12911323 style='border-top:none'>&nbsp;</td>                                                                   \n");
                    htmlStr.Append("							  <td colspan=2 class=xl13311323 style='border-left:none'>&nbsp;" + dt_d.Rows[dtR][3] + "</td>                                           \n");
                    htmlStr.Append("							  <td class=xl12811323 style='border-top:none'>&nbsp;</td>                                                                   \n");
                    htmlStr.Append("							  <td colspan=2 class=xl15511323 width=132 style='border-left:none;width:99pt'>&nbsp;" + dt_d.Rows[dtR][4] + "</td>                      \n");
                    htmlStr.Append("							  <td class=xl12811323 style='border-top:none'>&nbsp;</td>                                                                   \n");
                    htmlStr.Append("							 </tr>                                                                                                                       \n");
                }
            }
            if (v_count < 11)
            {
                for (int i = 0; i < 11 - v_count; i++)
                {
                    if (i == (11 - v_count) - 1)
                    {
                        htmlStr.Append("										 <tr class=xl7611323 height=25 style='mso-height-source:userset;height:" + heigh + "'>                                   \n");
                        htmlStr.Append("										  <td colspan=2 height=25 class=xl16111323 width=39 style='border-right:1pt solid black;                        \n");
                        htmlStr.Append("										  height:" + heigh + ";width:29pt'>&nbsp;</td>                                                                           \n");
                        htmlStr.Append("										  <td class=xl14111323 width=69 style='border-top:none;width:52pt;text-align:left' colspan=5>&nbsp;" + dt.Rows[0]["exchangerate"] + " &nbsp;</ td >\n");
                        htmlStr.Append("										          <td class=xl14211323 width=78 style='border-top:none;width:72.5‬pt'>&nbsp;</td>                                   \n");
                        htmlStr.Append("										  <td class=xl14311323 width=56 style='border-top:none;border-left:none;                                         \n");
                        htmlStr.Append("										  width:52.5pt'>&nbsp;</td>                                                                                        \n");
                        htmlStr.Append("										  <td class=xl14911323 width=30 style='border-top:none;width:28.75pt'>&nbsp;</td>                                   \n");
                        htmlStr.Append("										  <td class=xl14411323 width=6 style='border-top:none;width:5pt'>&nbsp;</td>                                     \n");
                        htmlStr.Append("										  <td colspan=2 class=xl15911323 style='border-left:none'>&nbsp;</td>                                            \n");
                        htmlStr.Append("										  <td class=xl14511323 width=6 style='border-top:none;width:5pt'>&nbsp;</td>                                     \n");
                        htmlStr.Append("										  <td colspan=2 class=xl15911323 style='border-left:none'>&nbsp;</td>                                            \n");
                        htmlStr.Append("										  <td class=xl13011323 width=6 style='border-top:none;width:5pt'>&nbsp;</td>                                     \n");
                        htmlStr.Append("										</tr>                                                                                                            \n");
                    }
                    else
                    {
                        htmlStr.Append("										<tr class=xl8411323 height=25 style='mso-height-source:userset;height:" + heigh + "'>                                    \n");
                        htmlStr.Append("										  <td colspan=2 height=25 class=xl15711323 width=39 style='border-right:1pt solid black;                        \n");
                        htmlStr.Append("										  height:" + heigh + ";width:29pt'>&nbsp;</td>                                                                           \n");
                        htmlStr.Append("										  <td colspan=5 class=xl13811323 width=288 style='border-right:1pt solid black; border-top:1pt dotted  windowtext;                                 \n");
                        htmlStr.Append("										  border-left:none;width:216pt'>&nbsp;</td>                                                                      \n");
                        htmlStr.Append("										  <td class=xl13211323 style='border-top:none;border-left:none'>&nbsp;</td>                                      \n");
                        htmlStr.Append("										  <td class=xl13311323 style='border-top:none;border-left:none'>&nbsp;</td>                                      \n");
                        htmlStr.Append("										  <td class=xl13411323 style='border-top:none'>&nbsp;</td>                                                       \n");
                        htmlStr.Append("										  <td class=xl12911323 style='border-top:none'>&nbsp;</td>                                                       \n");
                        htmlStr.Append("										  <td class=xl13311323 style='border-top:none;border-left:none'>&nbsp;</td>                                      \n");
                        htmlStr.Append("										  <td class=xl13411323 style='border-top:none'>&nbsp;</td>                                                       \n");
                        htmlStr.Append("										  <td class=xl12811323 style='border-top:none'>&nbsp;</td>                                                       \n");
                        htmlStr.Append("										  <td colspan=2 class=xl15511323 width=132 style='border-left:none;width:99pt'>&nbsp;</td>                       \n");
                        htmlStr.Append("										  <td class=xl12811323 style='border-top:none'>&nbsp;</td>                                                       \n");
                        htmlStr.Append("										 </tr>                                                                                                           \n");
                    }
                }
            }

            htmlStr.Append(" <tr class=xl7611323 height=26 style='mso-height-source:userset;height:25.6pt'>                                                                          \n");
            htmlStr.Append("  <td height=26 class=xl10311323 width=6 style='height:25.6pt;width:5pt'>&nbsp;</td>                                                                     \n");
            htmlStr.Append("  <td class=xl14611323 width=33 style='width:31.25‬pt'>&nbsp;</td>                                                                                           \n");
            htmlStr.Append("  <td class=xl14611323 width=69 style='width:52pt'>&nbsp;</td>                                                                                           \n");
            htmlStr.Append("  <td class=xl14611323 width=54 style='width:51.25pt'>&nbsp;</td>                                                                                           \n");
            htmlStr.Append("  <td class=xl14611323 width=40 style='width:30pt'>&nbsp;</td>                                                                                           \n");
            htmlStr.Append("  <td class=xl14611323 width=98 style='width:73pt'>&nbsp;</td>                                                                                           \n");
            htmlStr.Append("  <td class=xl10411323 width=27 style='width:20pt'>&nbsp;</td>                                                                                           \n");
            htmlStr.Append("  <td colspan=7 class=xl18611323 width=272 style='border-left:none;width:203pt'>&nbsp; C&#7897;ng                                                        \n");
            htmlStr.Append("  ti&#7873;n hàng (<font class='font1211323'>Total amount</font><font                                                                                    \n");
            htmlStr.Append("  class='font911323'>) : </font></td>                                                                                                                    \n");
            htmlStr.Append("  <td colspan=2 class=xl17111323>" + dt.Rows[0]["netamount_display"] + "</td>                                                                                                    \n");
            htmlStr.Append("  <td class=xl13111323>&nbsp;</td>                                                                                                                       \n");
            htmlStr.Append(" </tr>                                                                                                                                                   \n");
            htmlStr.Append(" <tr class=xl7611323 height=26 style='mso-height-source:userset;height:25.6pt'>                                                                          \n");
            htmlStr.Append("  <td height=26 class=xl10311323 width=6 style='height:25.6pt;border-top:none;                                                                           \n");
            htmlStr.Append("  width:5pt'>&nbsp;</td>                                                                                                                                 \n");
            htmlStr.Append("  <td colspan=4 class=xl12511323 width=196 style='width:148pt'><span                                                                                     \n");
            htmlStr.Append("  style='mso-spacerun:yes'> </span>Thu&#7871; su&#7845;t GTGT (<font                                                                                     \n");
            htmlStr.Append("  class='font1211323'>VAT rate</font><font class='font911323'>) :</font></td>                                                                            \n");
            htmlStr.Append("  <td class=xl15111323 width=98 style='border-top:none;width:73pt'>" + dt.Rows[0]["TaxRate"] + "</td>                                                                 \n");
            htmlStr.Append("  <td class=xl14711323 width=27 style='border-top:none;width:20pt'>&nbsp;</td>                                                                           \n");
            htmlStr.Append("  <td colspan=7 class=xl18611323 width=272 style='border-left:none;width:203pt'>&nbsp; Ti&#7873;n                                                        \n");
            htmlStr.Append("  thu&#7871; GTGT (<font class='font1211323'>VAT</font><font class='font911323'>)                                                                        \n");
            htmlStr.Append("  :</font></td>                                                                                                                                          \n");
            htmlStr.Append("  <td colspan=2 class=xl17111323>" + dt.Rows[0]["vatamount_display"] + "</td>                                                                                                  \n");
            htmlStr.Append("  <td class=xl8811323 style='border-top:none'>&nbsp;</td>                                                                                                \n");
            htmlStr.Append(" </tr>                                                                                                                                                   \n");
            if (read_prive.Length < 68)
            {
                htmlStr.Append(" 			                                                                                                                                             \n");
                htmlStr.Append(" 			<tr class=xl7611323 height=26 style='mso-height-source:userset;height:25.6pt'>                                                               \n");
                htmlStr.Append("			  <td height=26 class=xl10311323 width=6 style='height:25.6pt;border-top:none;                                                               \n");
                htmlStr.Append("			  width:5pt'>&nbsp;</td>                                                                                                                     \n");
                htmlStr.Append("			  <td class=xl12511323 width=33 style='border-top:none;width:31.25‬pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("			  <td class=xl12511323 width=69 style='border-top:none;width:52pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("			  <td class=xl12511323 width=54 style='border-top:none;width:51.25pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("			  <td class=xl12511323 width=40 style='border-top:none;width:30pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("			  <td class=xl12511323 width=98 style='border-top:none;width:73pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("			  <td class=xl14711323 width=27 style='border-top:none;width:20pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("			  <td colspan=7 class=xl18611323 width=272 style='border-left:none;width:203pt'>&nbsp; T&#7893;ng                                            \n");
                htmlStr.Append("			  ti&#7873;n thanh toán (<font class='font1211323'> Total payment</font><font                                                                \n");
                htmlStr.Append("			  class='font911323'>) :</font></td>                                                                                                         \n");
                htmlStr.Append("			  <td colspan=2 class=xl17111323>" + dt.Rows[0]["totalamount_display"] + "</td>                                                                               \n");
                htmlStr.Append("			  <td class=xl8811323>&nbsp;</td>                                                                                                            \n");
                htmlStr.Append("			 </tr>                                                                                                                                       \n");
                htmlStr.Append("			 <tr class=xl7611323 height=24 style='mso-height-source:userset;height:22.5pt'>                                                              \n");
                htmlStr.Append("			  <td height=24 class=xl8711323 style='height:22.5pt;border-top:none'>&nbsp;</td>                                                            \n");
                htmlStr.Append("			  <td colspan=5 class=xl16911323>S&#7889; ti&#7873;n vi&#7871;t b&#7857;ng                                                                   \n");
                htmlStr.Append("			  ch&#7919; (<font class='font1211323'>In words</font><font class='font711323'>):                                                            \n");
                htmlStr.Append("			  </font></td>                                                                                                                               \n");
                htmlStr.Append("			  <td colspan=10 class=xl16911323><font class='font711323'> " + read_prive + "                                                                 \n");
                htmlStr.Append("			  </font></td>                                                                                                                               \n");
                htmlStr.Append("			  <td class=xl8911323 width=6 style='width:5pt'>&nbsp;</td>                                                                                  \n");
                htmlStr.Append("			 </tr>                                                                                                                                       \n");
                htmlStr.Append("			 <tr height=18 style='mso-height-source:userset;height:17.625pt'>                                                                              \n");
                htmlStr.Append("			  <td height=18 class=xl8311323 style='height:17.625pt'>&nbsp;</td>                                                                            \n");
                htmlStr.Append("			  <td class=xl9011323>&nbsp;</td>                                                                                                            \n");
                htmlStr.Append("			  <td class=xl9011323>&nbsp;</td>                                                                                                            \n");
                htmlStr.Append("			  <td class=xl9011323>&nbsp;</td>                                                                                                            \n");
                htmlStr.Append("			  <td class=xl9111323 width=40 style='width:30pt'>&nbsp;</td>                                                                                \n");
                htmlStr.Append("			  <td class=xl9111323 width=98 style='width:73pt'>&nbsp;</td>                                                                                \n");
                htmlStr.Append("			  <td class=xl9111323 width=27 style='width:20pt'>&nbsp;</td>                                                                                \n");
                htmlStr.Append("			  <td class=xl9111323 width=78 style='width:72.5‬pt'>&nbsp;</td>                                                                                \n");
                htmlStr.Append("			  <td class=xl9111323 width=56 style='width:52.5pt'>&nbsp;</td>                                                                                \n");
                htmlStr.Append("			  <td class=xl9111323 width=30 style='width:28.75pt'>&nbsp;</td>                                                                                \n");
                htmlStr.Append("			  <td class=xl9111323 width=6 style='width:5pt'>&nbsp;</td>                                                                                  \n");
                htmlStr.Append("			  <td class=xl9111323 width=49 style='width:46.25pt'>&nbsp;</td>                                                                                \n");
                htmlStr.Append("			  <td class=xl9111323 width=47 style='width:43.75pt'>&nbsp;</td>                                                                                \n");
                htmlStr.Append("			  <td class=xl9111323 width=6 style='width:5pt'>&nbsp;</td>                                                                                  \n");
                htmlStr.Append("			  <td class=xl9111323 width=54 style='width:51.25pt'>&nbsp;</td>                                                                                \n");
                htmlStr.Append("			  <td class=xl9111323 width=78 style='width:72.5‬pt'>&nbsp;</td>                                                                                \n");
                htmlStr.Append("			  <td class=xl9211323 width=6 style='width:5pt'>&nbsp;</td>                                                                                  \n");
                htmlStr.Append("			 </tr>                                                                                                                                       \n");
                htmlStr.Append(" 			                                                                                                                                             \n");
            }
            else
            {
                htmlStr.Append(" 	                                                                                                                                                     \n");
                htmlStr.Append(" 			<tr class=xl7611323 height=26 style='mso-height-source:userset;height:25.6pt'>                                                               \n");
                htmlStr.Append("			  <td height=26 class=xl10311323 width=6 style='height:25.6pt;border-top:none;                                                               \n");
                htmlStr.Append("			  width:5pt'>&nbsp;</td>                                                                                                                     \n");
                htmlStr.Append("			  <td class=xl12511323 width=33 style='border-top:none;width:31.25‬pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("			  <td class=xl12511323 width=69 style='border-top:none;width:52pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("			  <td class=xl12511323 width=54 style='border-top:none;width:51.25pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("			  <td class=xl12511323 width=40 style='border-top:none;width:30pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("			  <td class=xl12511323 width=98 style='border-top:none;width:73pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("			  <td class=xl14711323 width=27 style='border-top:none;width:20pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("			  <td colspan=7 class=xl18611323 width=272 style='border-left:none;width:203pt'>&nbsp; T&#7893;ng                                            \n");
                htmlStr.Append("			  ti&#7873;n thanh toán (<font class='font1211323'> Total payment</font><font                                                                \n");
                htmlStr.Append("			  class='font911323'>) :</font></td>                                                                                                         \n");
                htmlStr.Append("			  <td colspan=2 class=xl17111323>" + dt.Rows[0]["totalamount_display"] + "</td>                                                                               \n");
                htmlStr.Append("			  <td class=xl8811323>&nbsp;</td>                                                                                                            \n");
                htmlStr.Append("			 </tr>                                                                                                                                       \n");
                htmlStr.Append("			 <tr class=xl7611323 height=24 style='mso-height-source:userset;height:22.5pt'>                                                              \n");
                htmlStr.Append("			  <td height=24 class=xl8711323 style='height:22.5pt;border-top:none'>&nbsp;</td>                                                            \n");
                htmlStr.Append("			  <td colspan=5 class=xl16911323>S&#7889; ti&#7873;n vi&#7871;t b&#7857;ng                                                                   \n");
                htmlStr.Append("			  ch&#7919; (<font class='font1211323'>In words</font><font class='font711323'>):                                                            \n");
                htmlStr.Append("			  </font></td>                                                                                                                               \n");
                htmlStr.Append("			  <td colspan=10 class=xl16911323><font class='font711323'>" + read_prive + "                                                                  \n");
                htmlStr.Append("			  </font></td>                                                                                                                               \n");
                htmlStr.Append("			  <td class=xl8911323 width=6 style='width:5pt'>&nbsp;</td>                                                                                  \n");
                htmlStr.Append("			 </tr>                                                                                                                                       \n");
                htmlStr.Append(" 	                                                                                                                                                     \n");
                htmlStr.Append(" 	                                                                                                                                                     \n");
            }
            htmlStr.Append("                                                                                                                                                         \n");
            htmlStr.Append("  <tr class=xl8011323 height = 24 style='mso-height-source:userset;height:18.0pt'>                                                                        \n");
            htmlStr.Append("  <td height = 24 class=xl9411323 style = 'height:18.0pt' > &nbsp;</td>                                                                                   \n");
            htmlStr.Append("  <td colspan = 4 class=xl16811323 width = 196 style='width:148pt'>Người chuyển đổi</td>                                                                         \n");
            htmlStr.Append("  <td colspan = 3 class=xl16811323 width = 203 style='width:151pt'>Ng&#432;&#7901;i                                                                       \n");
            htmlStr.Append("  mua hàng(Buyer)</td>                                                                                                                                    \n");
            htmlStr.Append("  <td colspan = 8 class=xl16811323 width = 326 style='width:244pt'>Ng&#432;&#7901;i                                                                       \n");
            htmlStr.Append("  bán hàng(Seller)</td>                                                                                                                                   \n");
            htmlStr.Append("  <td class=xl9311323 width = 6 style='width:4pt'>&nbsp;</td>                                                                                             \n");
            htmlStr.Append(" </tr>                                                                                                                                                    \n");
            htmlStr.Append(" <tr class=xl7011323 height = 18 style='mso-height-source:userset;height:14.1pt'>                                                                         \n");
            htmlStr.Append("  <td height = 18 class=xl9511323 style = 'height:14.1pt' > &nbsp;</td>                                                                                   \n");
            htmlStr.Append("  <td colspan = 4 class=xl16611323 width = 196 style='width:148pt'>(Ký, ghi rõ                                                                            \n");
            htmlStr.Append("  h&#7885; tên)</td>                                                                                                                                      \n");
            htmlStr.Append("  <td colspan = 3 class=xl16611323 width = 203 style='width:151pt'>(Ký, ghi rõ                                                                            \n");
            htmlStr.Append("  h&#7885; tên)</td>                                                                                                                                      \n");
            htmlStr.Append("  <td colspan = 8 class=xl16611323 width = 326 style='width:244pt'>(Ký, &#273;óng                                                                         \n");
            htmlStr.Append("  d&#7845;u, ghi rõ h&#7885; tên)</td>                                                                                                                    \n");
            htmlStr.Append("  <td class=xl9611323 width = 6 style='width:4pt'>&nbsp;</td>                                                                                             \n");
            htmlStr.Append(" </tr>                                                                                                                                                    \n");
            htmlStr.Append(" <tr class=xl9911323 height = 21 style='mso-height-source:userset;height:15.6pt'>                                                                         \n");
            htmlStr.Append("  <td height = 21 class=xl9711323 style = 'height:15.6pt' > &nbsp;</td>                                                                                   \n");
            htmlStr.Append("  <td colspan = 4 class=xl16711323 width = 196 style='width:148pt'>(Signature &amp;                                                                       \n");
            htmlStr.Append("  full name)</td>                                                                                                                                         \n");
            htmlStr.Append("  <td colspan = 3 class=xl16711323 width = 203 style='width:151pt'>(Signature &amp;                                                                       \n");
            htmlStr.Append("  full name)</td>                                                                                                                                         \n");
            htmlStr.Append("  <td colspan = 8 class=xl16711323 width = 326 style='width:244pt'>(Signature,                                                                            \n");
            htmlStr.Append("  stamp &amp; full name)</td>                                                                                                                             \n");
            htmlStr.Append("  <td class=xl9811323 width = 6 style='width:4pt'>&nbsp;</td>                                                                                             \n");
            htmlStr.Append(" </tr>                                                                                                                                                    \n");
            htmlStr.Append("                                                                                                                                                         \n");
            if (dt_d.Rows[0][0].ToString().Length < 28)
            {
                htmlStr.Append(" <tr height=21 style='mso-height-source:userset;height:10pt'>                                                                                         \n");
                htmlStr.Append("  <td height=21 class=xl12611323 style='height:10pt'>&nbsp;</td>                                                                                      \n");
                htmlStr.Append("  <td class=xl6611323>&nbsp;</td>                                                                                                                        \n");
                htmlStr.Append("  <td class=xl7111323 colspan=15>&nbsp;</td>                                                                                                                        \n");
                htmlStr.Append(" </tr>                                                                                                                                                   \n");
            }


            htmlStr.Append(" <tr height=22 style='height:18.75pt'>                                                                                                                      \n");
            htmlStr.Append("  <td height=22 class=xl12611323 style='height:18.75pt'>&nbsp;</td>                                                                                         \n");
            htmlStr.Append("  <td class=xl6611323>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append("  <td class=xl6611323>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append("  <td class=xl6611323>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append("  <td class=xl7211323 width=40 style='width:30pt'>&nbsp;</td>                                                                                            \n");
            htmlStr.Append("  <td class=xl7211323 width=98 style='width:73pt'>&nbsp;</td>                                                                                            \n");
            htmlStr.Append("  <td class=xl7211323 width=27 style='width:20pt'>&nbsp;</td>                                                                                            \n");
            htmlStr.Append("  <td class=xl7211323 width=78 style='width:72.5‬pt'>&nbsp;</td>                                                                                            \n");
            htmlStr.Append("  <td class=xl10711323 colspan=4>Signature Valid</td>                                                                                                    \n");


            htmlStr.Append("  <td align=left valign=top><![if !vml]><span style='mso-ignore:vglayout;                                                                                \n");
            htmlStr.Append("  position:absolute;z-index:1;margin-left:18px;margin-top:7px;width:79px;                                                                                \n");
            htmlStr.Append("  height:56px'><img width=79 height=56                                                                                                                   \n");
            htmlStr.Append("  src='D:\\webproject\\e-invoice-ws\\02.Web\\EInvoice\\img\\check_signed.png'                                                                                   \n");
            htmlStr.Append("  v:shapes='Picture_x0020_8'></span><![endif]><span style='mso-ignore:vglayout2'>                                                                        \n");
            htmlStr.Append("  <table cellpadding=0 cellspacing=0>                                                                                                                    \n");
            htmlStr.Append("   <tr>                                                                                                                                                  \n");
            htmlStr.Append("    <td height=22 class=xl10811323 width=47 style='height:16.8pt;width:43.75pt'>&nbsp;</td>                                                                 \n");
            htmlStr.Append("   </tr>                                                                                                                                                 \n");
            htmlStr.Append("  </table>                                                                                                                                               \n");
            htmlStr.Append("  </span></td>                                                                                                                                           \n");

            htmlStr.Append("  <td class=xl10911323>&nbsp;</td>                                                                                                                       \n");
            htmlStr.Append("  <td class=xl10911323>&nbsp;</td>                                                                                                                       \n");
            htmlStr.Append("  <td class=xl11011323>&nbsp;</td>                                                                                                                       \n");
            htmlStr.Append("  <td class=xl11211323>&nbsp;</td>                                                                                                                       \n");
            htmlStr.Append(" </tr>                                                                                                                                                   \n");
            htmlStr.Append(" <tr class=xl11811323 height=21 style='mso-height-source:userset;height:18.75pt'>                                                                           \n");
            htmlStr.Append("  <td height=21 class=xl11711323 style='height:18.75pt'>&nbsp;</td>                                                                                         \n");
            htmlStr.Append("  <td class=xl11811323>&nbsp;</td>                                                                                                                       \n");
            htmlStr.Append("  <td class=xl11811323>&nbsp;</td>                                                                                                                       \n");
            htmlStr.Append("  <td class=xl11811323>&nbsp;</td>                                                                                                                       \n");
            htmlStr.Append("  <td class=xl11911323>&nbsp;</td>                                                                                                                       \n");
            htmlStr.Append("  <td class=xl11911323>&nbsp;</td>                                                                                                                       \n");
            htmlStr.Append("  <td class=xl11911323>&nbsp;</td>                                                                                                                       \n");
            htmlStr.Append("  <td class=xl11911323>&nbsp;</td>                                                                                                                       \n");
            htmlStr.Append("  <td colspan=8 class=xl16311323 style='border-right:1pt solid black'><font                                                                             \n");
            htmlStr.Append("  class='font1511323'>&#272;&#432;&#7907;c ký b&#7903;i:</font><font                                                                                     \n");
            htmlStr.Append("  class='font1611323'> </font><font class='font1711323'style='font-size:11pt'>" + dt.Rows[0]["SignedBy"] + "</font></td>                                                                   \n");
            htmlStr.Append("  <td class=xl12011323>&nbsp;</td>                                                                                                                       \n");
            htmlStr.Append(" </tr>                                                                                                                                                   \n");
            htmlStr.Append(" <tr height=21 style='height:18.75pt'>                                                                                                                      \n");
            htmlStr.Append("  <td height=21 class=xl12611323 style='height:18.75pt'>&nbsp;</td>                                                                                         \n");
            htmlStr.Append("  <td class=xl6511323></td>                                                                                                                              \n");
            htmlStr.Append("  <td class=xl6611323>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append("  <td class=xl6611323>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append("  <td class=xl7211323 width=40 style='width:30pt'>&nbsp;</td>                                                                                            \n");
            htmlStr.Append("  <td class=xl7211323 width=98 style='width:73pt'>&nbsp;</td>                                                                                            \n");
            htmlStr.Append("  <td class=xl7211323 width=27 style='width:20pt'>&nbsp;</td>                                                                                            \n");
            htmlStr.Append("  <td class=xl7211323 width=78 style='width:72.5‬pt'>&nbsp;</td>                                                                                            \n");
            htmlStr.Append("  <td class=xl11311323 colspan=4>Ngày ký: <font class='font1811323'>" + dt.Rows[0]["SignedDate"] + "</font></td>                                                           \n");
            htmlStr.Append("  <td class=xl11411323>&nbsp;</td>                                                                                                                       \n");
            htmlStr.Append("  <td class=xl11411323>&nbsp;</td>                                                                                                                       \n");
            htmlStr.Append("  <td class=xl11411323>&nbsp;</td>                                                                                                                       \n");
            htmlStr.Append("  <td class=xl11511323>&nbsp;</td>                                                                                                                       \n");
            htmlStr.Append("  <td class=xl11111323>&nbsp;</td>                                                                                                                       \n");
            htmlStr.Append(" </tr>                                                                                                                                                   \n");
            htmlStr.Append(" <tr height=18 style='mso-height-source:userset;height:17.625pt'>                                                                                          \n");
            htmlStr.Append("  <td height=18 class=xl12611323 style='height:17.625pt'>&nbsp;</td>                                                                                       \n");
            htmlStr.Append("  <td class=xl11611323 colspan=3>Mã nhận hóa &#273;&#417;n:</td>                                                                                         \n");
            htmlStr.Append("  <td class=xl7211323 width=40 style='width:30pt' colspan=12>&nbsp;" + dt.Rows[0]["matracuu"] + "</td>                                                   \n");
            htmlStr.Append("  <td class=xl10011323 width=6 style='width:5pt'>&nbsp;</td>                                                                                             \n");
            htmlStr.Append(" </tr>                                                                                                                                                   \n");
            htmlStr.Append(" <tr height=20 style='mso-height-source:userset;height:18.75pt'>                                                                                          \n");
            htmlStr.Append("  <td height=20 class=xl12611323 style='height:18.75pt'>&nbsp;</td>                                                                                       \n");
            htmlStr.Append("  <td class=xl6611323 colspan=7>Tra c&#7913;u t&#7841;i Website: <font                                                                                   \n");
            htmlStr.Append("  class='font611323'><span style='mso-spacerun:yes'> </span></font><font                                                                                 \n");
            htmlStr.Append("  class='font1911323'>" + dt.Rows[0]["WEBSITE_EI"] + "</font></td>                                                                                   \n");
            htmlStr.Append("  <td class=xl7211323 width=56 style='width:52.5pt'>&nbsp;</td>                                                                                            \n");
            htmlStr.Append("  <td class=xl7211323 width=30 style='width:28.75pt'>&nbsp;</td>                                                                                            \n");
            htmlStr.Append("  <td class=xl7211323 width=6 style='width:5pt'>&nbsp;</td>                                                                                              \n");
            htmlStr.Append("  <td class=xl7211323 width=49 style='width:46.25pt'>&nbsp;</td>                                                                                            \n");
            htmlStr.Append("  <td class=xl7211323 width=47 style='width:43.75pt'>&nbsp;</td>                                                                                            \n");
            htmlStr.Append("  <td class=xl7211323 width=6 style='width:5pt'>&nbsp;</td>                                                                                              \n");
            htmlStr.Append("  <td class=xl7211323 width=54 style='width:51.25pt'>&nbsp;</td>                                                                                            \n");
            htmlStr.Append("  <td class=xl7211323 width=78 style='width:72.5‬pt'>&nbsp;</td>                                                                                            \n");
            htmlStr.Append("  <td class=xl10011323 width=6 style='width:5pt'>&nbsp;</td>                                                                                             \n");
            htmlStr.Append(" </tr>                                                                                                                                                   \n");
            htmlStr.Append(" <tr height=16 style='mso-height-source:userset;height:10.75pt'>                                                                                          \n");
            htmlStr.Append("  <td height=16 class=xl12611323 style='height:10.75pt'>&nbsp;</td>                                                                                       \n");
            htmlStr.Append("  <td colspan=15 class=xl17011323>(C&#7847;n ki&#7875;m tra, &#273;&#7889;i                                                                              \n");
            htmlStr.Append("  chi&#7871;u khi l&#7853;p, giao nh&#7853;n hóa &#273;&#417;n)</td>                                                                                     \n");
            htmlStr.Append("  <td class=xl10011323 width=6 style='width:5pt'>&nbsp;</td>                                                                                             \n");
            htmlStr.Append(" </tr>                                                                                                                                                   \n");
            htmlStr.Append(" <tr height=16 style='mso-height-source:userset;height:10.75pt'>                                                                                          \n");
            htmlStr.Append("  <td height=16 class=xl8311323 style='height:10.75pt'>&nbsp;</td>                                                                                        \n");
            htmlStr.Append("  <td colspan=15 class=xl17511323>" + dt.Rows[0]["CONTRACT_INFO_EI"] + "</td>                                                                          \n");
            htmlStr.Append("  <td class=xl10211323 width=6 style='width:5pt'>&nbsp;</td>                                                                                             \n");
            htmlStr.Append(" </tr>                                                                                                                                                   \n");
            htmlStr.Append(" <![if supportMisalignedColumns]>                                                                                                                        \n");
            htmlStr.Append(" <tr height=0 style='display:none'>                                                                                                                      \n");
            htmlStr.Append("  <td width=6 style='width:5pt'></td>                                                                                                                    \n");
            htmlStr.Append("  <td width=33 style='width:31.25‬pt'></td>                                                                                                                  \n");
            htmlStr.Append("  <td width=69 style='width:52pt'></td>                                                                                                                  \n");
            htmlStr.Append("  <td width=54 style='width:51.25pt'></td>                                                                                                                  \n");
            htmlStr.Append("  <td width=40 style='width:30pt'></td>                                                                                                                  \n");
            htmlStr.Append("  <td width=98 style='width:73pt'></td>                                                                                                                  \n");
            htmlStr.Append("  <td width=27 style='width:20pt'></td>                                                                                                                  \n");
            htmlStr.Append("  <td width=78 style='width:72.5‬pt'></td>                                                                                                                  \n");
            htmlStr.Append("  <td width=56 style='width:52.5pt'></td>                                                                                                                  \n");
            htmlStr.Append("  <td width=30 style='width:28.75pt'></td>                                                                                                                  \n");
            htmlStr.Append("  <td width=6 style='width:5pt'></td>                                                                                                                    \n");
            htmlStr.Append("  <td width=49 style='width:46.25pt'></td>                                                                                                                  \n");
            htmlStr.Append("  <td width=47 style='width:43.75pt'></td>                                                                                                                  \n");
            htmlStr.Append("  <td width=6 style='width:5pt'></td>                                                                                                                    \n");
            htmlStr.Append("  <td width=54 style='width:51.25pt'></td>                                                                                                                  \n");
            htmlStr.Append("  <td width=78 style='width:72.5‬pt'></td>                                                                                                                  \n");
            htmlStr.Append("  <td width=6 style='width:5pt'></td>                                                                                                                    \n");
            htmlStr.Append(" </tr>                                                                                                                                                   \n");
            htmlStr.Append(" <![endif]>                                                                                                                                              \n");
            htmlStr.Append("</table>                                                                                                                                                 \n");
            htmlStr.Append("</body>                                                                                                                                                  \n");
            htmlStr.Append("</html>																																					 \n");

            // insert xml in database 

            //string filePath = "C:\\Users\\genuwin\\Desktop\\" + tei_einvoice_m_pk + ".xml";
            string filePath = "D:\\webproject\\e-invoice-ws\\02.Web\\AttachFileXml\\" + tei_einvoice_m_pk + ".xml";
            if (File.Exists(filePath))
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

            doc.Save(filePath);

            //}


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
            tempLob.EndChunkWrite();// EndBatch();;

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
    }
}