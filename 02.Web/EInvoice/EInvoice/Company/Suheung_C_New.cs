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
    public class Suheung_C_New
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


            htmlStr.Append("<html xmlns:v='urn:schemas-microsoft-com:vml'																										\n");
            htmlStr.Append("xmlns:o='urn:schemas-microsoft-com:office:office'                                                                                                  \n");
            htmlStr.Append("xmlns:x='urn:schemas-microsoft-com:office:excel'                                                                                                   \n");
            htmlStr.Append("xmlns='http://www.w3.org/TR/REC-html40'>                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                   \n");
            htmlStr.Append("<head>                                                                                                                                             \n");
            htmlStr.Append("<meta charset='UTF-8'>\n");
            htmlStr.Append("<meta http-equiv=Content-Type content='text/html; charset=UTF-8'>                                                                           \n");
            htmlStr.Append("<meta name=ProgId content=Excel.Sheet>                                                                                                             \n");
            htmlStr.Append("<meta name=Generator content='Microsoft Excel 15'>                                                                                                 \n");
            htmlStr.Append("<link rel=File-List href='hóa%20&#273;&#417;n%20m&#7851;u%20suheung_files/filelist.xml'>                                                           \n");
            htmlStr.Append("<!--[if !mso]>                                                                                                                                     \n");
            htmlStr.Append("<style>                                                                                                                                            \n");
            htmlStr.Append("v\\:* {behavior:url(#default#VML);}                                                                                                                 \n");
            htmlStr.Append("o\\:* {behavior:url(#default#VML);}                                                                                                                 \n");
            htmlStr.Append("x\\:* {behavior:url(#default#VML);}                                                                                                                 \n");
            htmlStr.Append(".shape {behavior:url(#default#VML);}                                                                                                               \n");
            htmlStr.Append("</style>                                                                                                                                           \n");
            htmlStr.Append("<![endif]-->                                                                                                                                       \n");
            htmlStr.Append("<style id='hóa &#273;&#417;n m&#7851;u suheung_15939_Styles'>                                                                                      \n");
            htmlStr.Append("<!--table                                                                                                                                          \n");
            htmlStr.Append("	{mso-displayed-decimal-separator:'\\.';                                                                                                          \n");
            htmlStr.Append("	mso-displayed-thousand-separator:'\\,';}                                                                                                         \n");
            htmlStr.Append(".font515939                                                                                                                                        \n");
            htmlStr.Append("	{color:windowtext;                                                                                                                              \n");
            htmlStr.Append("	font-size:14pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                            \n");
            htmlStr.Append("	.font1226793                                                                                                                                     \n");
            htmlStr.Append("	 {                                                                                                                                               \n");
            htmlStr.Append("	    color: blue;                                                                                                                                 \n");
            htmlStr.Append("	    font-size:12.5pt;                                                                                                                          \n");
            htmlStr.Append("	    font-weight:700;                                                                                                                           \n");
            htmlStr.Append("	    text-decoration:none;                                                                                                                      \n");
            htmlStr.Append("	    font-family:'Times New Roman', serif;                                                                                                      \n");
            htmlStr.Append("	    mso-font-charset:0;                                                                                                                      \n");
            htmlStr.Append("	}                                                                                                                                                \n");
            htmlStr.Append(".font615939                                                                                                                                        \n");
            htmlStr.Append("	{color:windowtext;                                                                                                                              \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                            \n");
            htmlStr.Append(".font715939                                                                                                                                        \n");
            htmlStr.Append("	{color:windowtext;                                                                                                                              \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                            \n");
            htmlStr.Append(".font815939                                                                                                                                        \n");
            htmlStr.Append("	{color:windowtext;                                                                                                                              \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                            \n");
            htmlStr.Append(".font915939                                                                                                                                        \n");
            htmlStr.Append("	{color:windowtext;                                                                                                                              \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                            \n");
            htmlStr.Append(".font1015939                                                                                                                                       \n");
            htmlStr.Append("	{color:windowtext;                                                                                                                              \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                            \n");
            htmlStr.Append(".font1115939                                                                                                                                       \n");
            htmlStr.Append("	{color:#993300;                                                                                                                                 \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                            \n");
            htmlStr.Append(".font1215939                                                                                                                                       \n");
            htmlStr.Append("	{color:red;                                                                                                                                     \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                            \n");
            htmlStr.Append(".font1315939                                                                                                                                       \n");
            htmlStr.Append("	{color:red;                                                                                                                                     \n");
            htmlStr.Append("	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                            \n");
            htmlStr.Append(".font1415939                                                                                                                                       \n");
            htmlStr.Append("	{color:red;                                                                                                                                     \n");
            htmlStr.Append("	font-size:8.75‬pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                            \n");
            htmlStr.Append(".font1515939                                                                                                                                       \n");
            htmlStr.Append("	{color:red;                                                                                                                                     \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                            \n");
            htmlStr.Append(".xl6515939                                                                                                                                         \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:red;                                                                                                                                      \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:bottom;                                                                                                                          \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl6615939                                                                                                                                         \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:bottom;                                                                                                                          \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl6715939                                                                                                                                         \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:bottom;                                                                                                                          \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl6815939                                                                                                                                         \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:bottom;                                                                                                                          \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl6915939                                                                                                                                         \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:red;                                                                                                                                      \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl7015939                                                                                                                                         \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:red;                                                                                                                                      \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl7115939                                                                                                                                         \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:red;                                                                                                                                      \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl7215939                                                                                                                                         \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl7315939                                                                                                                                         \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl7415939                                                                                                                                         \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl7515939                                                                                                                                         \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl7615939                                                                                                                                         \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl7715939                                                                                                                                         \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:16.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl7815939                                                                                                                                         \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:22.5pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl7915939                                                                                                                                         \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:16.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl8015939                                                                                                                                         \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl8115939                                                                                                                                         \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl8215939                                                                                                                                         \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:14pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl8315939                                                                                                                                         \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl8415939                                                                                                                                         \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl8515939                                                                                                                                         \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:14.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl8615939                                                                                                                                         \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:14.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl8715939                                                                                                                                         \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:14pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl8815939                                                                                                                                         \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl8915939                                                                                                                                         \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                         \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl9015939                                                                                                                                         \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:14pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl9115939                                                                                                                                         \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:14pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl9215939                                                                                                                                         \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append(".xl9315939                                                                                                                                         \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append(".xl9415939                                                                                                                                         \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl9515939                                                                                                                                         \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:right;                                                                                                                               \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl9615939                                                                                                                                         \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl9715939                                                                                                                                         \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:14pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl9815939                                                                                                                                         \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl9915939                                                                                                                                         \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:right;                                                                                                                               \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl10015939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl10115939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:14pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl10215939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:14pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl10315939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl10415939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl10515939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl10615939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append(".xl10715939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append(".xl10815939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:'_\\(* \\#\\,\\#\\#0_\\)\\;_\\(* \\\\(\\#\\,\\#\\#0\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                 \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append(".xl10915939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:'_\\(* \\#\\,\\#\\#0_\\)\\;_\\(* \\\\(\\#\\,\\#\\#0\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                         \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append(".xl11015939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:'_\\(* \\#\\,\\#\\#0_\\)\\;_\\(* \\\\(\\#\\,\\#\\#0\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                         \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append(".xl11115939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:'_\\(* \\#\\,\\#\\#0_\\)\\;_\\(* \\\\(\\#\\,\\#\\#0\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                         \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append(".xl11215939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:14pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl11315939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl11415939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append(".xl11515939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                         \n");
            htmlStr.Append("	text-align:right;                                                                                                                               \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append(".xl11615939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                         \n");
            htmlStr.Append("	text-align:right;                                                                                                                               \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append(".xl11715939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                         \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append(".xl11815939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                         \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append(".xl11915939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                         \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append(".xl12015939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                         \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append(".xl12115939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                         \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl12215939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                         \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl12315939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                         \n");
            htmlStr.Append("	text-align:right;                                                                                                                               \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl12415939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                         \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl12515939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                         \n");
            htmlStr.Append("	text-align:left;                                                                                                                                \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl12615939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:14pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl12715939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl12815939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:14pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl12915939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:14pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                         \n");
            htmlStr.Append("	text-align:right;                                                                                                                               \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append(".xl13015939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:14pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl13115939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:14pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                         \n");
            htmlStr.Append("	text-align:right;                                                                                                                               \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append(".xl13215939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl13315939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl13415939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl13515939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl13615939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl13715939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl13815939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl13915939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:14pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl14015939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:14pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl14115939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl14215939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl14315939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl14415939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl14515939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl14615939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:14pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl14715939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:14pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl14815939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl14915939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:14pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl15015939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:#C00000;                                                                                                                                  \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl15115939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:14pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl15215939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	mso-background-source:auto;                                                                                                                     \n");
            htmlStr.Append("	mso-pattern:auto;                                                                                                                               \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl15315939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:red;                                                                                                                                      \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:general;                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append(".xl15415939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:22.5pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl15515939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:16.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl15615939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:14pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl156159391                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:14pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                                \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl15715939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append(".xl15815939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append(".xl15915939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:14pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl16015939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:14pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl16115939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl16215939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl16315939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl16415939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl16515939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:14pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl16615939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl16715939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl16815939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                         \n");
            htmlStr.Append("	text-align:right;                                                                                                                               \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append(".xl16915939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            // htmlStr.Append("	color:red;                                                                                                                                      \n");
            htmlStr.Append("	font-size:14pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl17015939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:red;                                                                                                                                      \n");
            htmlStr.Append("	font-size:20.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl17115939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:red;                                                                                                                                      \n");
            htmlStr.Append("	font-size:20.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl17215939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:red;                                                                                                                                      \n");
            htmlStr.Append("	font-size:20.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl17315939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:red;                                                                                                                                      \n");
            htmlStr.Append("	font-size:20.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl17415939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:red;                                                                                                                                      \n");
            htmlStr.Append("	font-size:20.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl17515939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl17615939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                         \n");
            htmlStr.Append("	text-align:right;                                                                                                                               \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append(".xl17715939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                         \n");
            htmlStr.Append("	text-align:right;                                                                                                                               \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append(".xl17815939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl17915939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl18015939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                         \n");
            htmlStr.Append("	text-align:right;                                                                                                                               \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append(".xl18115939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                         \n");
            htmlStr.Append("	text-align:right;                                                                                                                               \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append(".xl18215939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:14pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl18315939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:14pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl18415939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl18515939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:9.0pt;                                                                                                                                \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl18615939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:red;                                                                                                                                      \n");
            htmlStr.Append("	font-size:8.75‬pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:none;                                                                                                                              \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append(".xl18715939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:red;                                                                                                                                      \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append(".xl18815939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:red;                                                                                                                                      \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("	border-left:none;                                                                                                                               \n");
            htmlStr.Append("	background:white;                                                                                                                               \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append(".xl18915939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	mso-background-source:auto;                                                                                                                     \n");
            htmlStr.Append("	mso-pattern:auto;                                                                                                                               \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append(".xl19015939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	mso-background-source:auto;                                                                                                                     \n");
            htmlStr.Append("	mso-pattern:auto;                                                                                                                               \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl19115939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:14pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	mso-background-source:auto;                                                                                                                     \n");
            htmlStr.Append("	mso-pattern:auto;                                                                                                                               \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl19215939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	mso-background-source:auto;                                                                                                                     \n");
            htmlStr.Append("	mso-pattern:auto;                                                                                                                               \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append(".xl19315939                                                                                                                                        \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("	font-size:14pt;                                                                                                                               \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                              \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("	mso-background-source:auto;                                                                                                                     \n");
            htmlStr.Append("	mso-pattern:auto;                                                                                                                               \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("	.xl18720842      \n");
            htmlStr.Append("	{                    \n");
            htmlStr.Append("	    padding: 0px;    \n");
            htmlStr.Append("	    mso-ignore:padding;    \n");
            htmlStr.Append("	    color: windowtext;   \n");
            htmlStr.Append("	    font-size:14pt;  \n");
            htmlStr.Append("	    font-weight:700;   \n");
            htmlStr.Append("	    font-style:normal;     \n");
            htmlStr.Append("	    text-decoration:none;  \n");
            htmlStr.Append("	    font-family:'Times New Roman', serif;  \n");
            htmlStr.Append("	    mso-font-charset:0;  \n");
            htmlStr.Append("	    mso-number-format:General;   \n");
            htmlStr.Append("	    text-align:center;     \n");
            htmlStr.Append("	    vertical-align:middle;     \n");
            htmlStr.Append("	    border-top:.5pt solid windowtext;  \n");
            htmlStr.Append("	    border-right:none; \n");
            htmlStr.Append("	    border-bottom:none;    \n");
            htmlStr.Append("	    border-left:.5pt solid windowtext; \n");
            htmlStr.Append("	    background: white; \n");
            htmlStr.Append("	    mso-pattern:black none; \n");
            htmlStr.Append("	    white-space:normal; \n");
            htmlStr.Append("	} \n");
            htmlStr.Append(".xl12520842\n");
            htmlStr.Append("{\n");
            htmlStr.Append("    padding: 0px;\n");
            htmlStr.Append("    mso-ignore:padding;\n");
            htmlStr.Append("    color: windowtext;\n");
            htmlStr.Append("    font-size:13.0pt;\n");
            htmlStr.Append("    font-weight:400;\n");
            htmlStr.Append("    font-style:normal;\n");
            htmlStr.Append("    text-decoration:none;\n");
            htmlStr.Append("    font-family:'Times New Roman', serif;\n");
            htmlStr.Append("    mso-font-charset:0;\n");
            htmlStr.Append("    mso-number-format:'\\@';\n");
            htmlStr.Append("    text-align:left;\n");
            htmlStr.Append("    vertical-align:middle;\n");
            htmlStr.Append("    border-top:none\n");
            htmlStr.Append("    border-right:.5pt solid windowtext;\n");
            htmlStr.Append("    border-bottom:none;\n");
            htmlStr.Append("    border-left:none;\n");
            htmlStr.Append("    background: white;\n");
            htmlStr.Append("    mso-pattern:black none;\n");
            htmlStr.Append("    white-space:normal;\n");
            htmlStr.Append("}\n");
            htmlStr.Append(".xl16520842\n");
            htmlStr.Append(" {\n");
            htmlStr.Append("    padding: 0px;\n");
            htmlStr.Append("    mso-ignore:padding;\n");
            htmlStr.Append("    color: windowtext;\n");
            htmlStr.Append("    font-size:13.0pt;\n");
            htmlStr.Append("    font-weight:400;\n");
            htmlStr.Append("    font-style:normal;\n");
            htmlStr.Append("    text-decoration:none;\n");
            htmlStr.Append("    font-family:'Times New Roman', serif;\n");
            htmlStr.Append("    mso-font-charset:0;\n");
            htmlStr.Append("    mso-number-format:General;\n");
            htmlStr.Append("    text-align:center;\n");
            htmlStr.Append("    vertical-align:middle;\n");
            htmlStr.Append("    border-top:none;\n");
            htmlStr.Append("    border-right:none;\n");
            htmlStr.Append("    border-bottom:none;\n");
            htmlStr.Append("    border-left:.5pt solid windowtext;\n");
            htmlStr.Append("    background: white;\n");
            htmlStr.Append("    mso-pattern:black none;\n");
            htmlStr.Append("    white-space:normal;\n");
            htmlStr.Append("}\n");
            htmlStr.Append(".xl10620842\n");
            htmlStr.Append("{\n");
            htmlStr.Append("    padding: 0px;\n");
            htmlStr.Append("    mso-ignore:padding;\n");
            htmlStr.Append("    color: windowtext;\n");
            htmlStr.Append("    font-size:13.0pt;\n");
            htmlStr.Append("    font-weight:400;\n");
            htmlStr.Append("    font-style:normal;\n");
            htmlStr.Append("    text-decoration:none;\n");
            htmlStr.Append("    font-family:'Times New Roman', serif;\n");
            htmlStr.Append("    mso-font-charset:0;\n");
            htmlStr.Append("    mso-number-format:General;\n");
            htmlStr.Append("    text-align:center;\n");
            htmlStr.Append("    vertical-align:middle;\n");
            htmlStr.Append("    border-top:none;\n");
            htmlStr.Append("    border-right:.5pt solid windowtext;\n");
            htmlStr.Append("    border-bottom:none;\n");
            htmlStr.Append("    border-left:.5pt solid windowtext;\n");
            htmlStr.Append("    background: white;\n");
            htmlStr.Append("    mso-pattern:black none\n");
            htmlStr.Append("    white-space:nowrap;\n");
            htmlStr.Append("    mso-text-control:shrinktofit;\n");
            htmlStr.Append("}\n");
            htmlStr.Append(" .xl12220842\n");
            htmlStr.Append("{\n");
            htmlStr.Append("    padding: 0px;\n");
            htmlStr.Append("    mso-ignore:padding;\n");
            htmlStr.Append("    color: windowtext;\n");
            htmlStr.Append("    font-size:13.0pt;\n");
            htmlStr.Append("    font-weight:400;\n");
            htmlStr.Append("    font-style:normal;\n");
            htmlStr.Append("    text-decoration:none;\n");
            htmlStr.Append("    font-family:'Times New Roman', serif;\n");
            htmlStr.Append("    mso-font-charset:0;\n");
            htmlStr.Append("    mso-number-format:'\\@';\n");
            htmlStr.Append("    text-align:right;\n");
            htmlStr.Append("    vertical-align:middle;\n");
            htmlStr.Append("    border-top:none;\n");
            htmlStr.Append("    border-right:none;\n");
            htmlStr.Append("    border-bottom:none;\n");
            htmlStr.Append("    border-left:.5pt solid windowtext;\n");
            htmlStr.Append("    background: white;\n");
            htmlStr.Append("    mso-pattern:black none;\n");
            htmlStr.Append("    white-space:normal;\n");
            htmlStr.Append("}\n");
            htmlStr.Append(".xl12120842\n");
            htmlStr.Append("{\n");
            htmlStr.Append("    padding: 0px;\n");
            htmlStr.Append("    mso-ignore:padding;\n");
            htmlStr.Append("    color: windowtext;\n");
            htmlStr.Append("    font-size:13.0pt;\n");
            htmlStr.Append("    font-weight:400;\n");
            htmlStr.Append("    font-style:normal;\n");
            htmlStr.Append("    text-decoration:none;\n");
            htmlStr.Append("    font-family:'Times New Roman', serif;\n");
            htmlStr.Append("    mso-font-charset:0;\n");
            htmlStr.Append("    mso-number-format:'\\@';\n");
            htmlStr.Append("    text-align:center;\n");
            htmlStr.Append("    vertical-align:middle;\n");
            htmlStr.Append("    border-top:none;\n");
            htmlStr.Append("    border-right:.5pt solid windowtext;\n");
            htmlStr.Append("    border-bottom:none;\n");
            htmlStr.Append("    border-left:.5pt solid windowtext;\n");
            htmlStr.Append("    background: white;\n");
            htmlStr.Append("    mso-pattern:black none;\n");
            htmlStr.Append("    white-space:normal;\n");
            htmlStr.Append("}\n");
            htmlStr.Append(".xl10720842\n");
            htmlStr.Append("{\n");
            htmlStr.Append("    padding: 0px;\n");
            htmlStr.Append("    mso-ignore:padding;\n");
            htmlStr.Append("    color: windowtext;\n");
            htmlStr.Append("    font-size:13.0pt;\n");
            htmlStr.Append("    font-weight:400;\n");
            htmlStr.Append("    font-style:normal;\n");
            htmlStr.Append("    text-decoration:none;\n");
            htmlStr.Append("    font-family:'Times New Roman', serif;\n");
            htmlStr.Append("    mso-font-charset:0;\n");
            htmlStr.Append("    mso-number-format:General;\n");
            htmlStr.Append("    text-align:center;\n");
            htmlStr.Append("    vertical-align:middle;\n");
            htmlStr.Append("    border-top:none;\n");
            htmlStr.Append("    border-right:none;\n");
            htmlStr.Append("    border-bottom:none;\n");
            htmlStr.Append("    border-left:.5pt solid windowtext;\n");
            htmlStr.Append("    background: white;\n");
            htmlStr.Append("    mso-pattern:black none;\n");
            htmlStr.Append("    white-space:nowrap;\n");
            htmlStr.Append("    mso-text-control:shrinktofit;\n");
            htmlStr.Append("}\n");
            htmlStr.Append(".xl11220842\n");
            htmlStr.Append("{\n");
            htmlStr.Append("    padding: 0px;\n");
            htmlStr.Append("    mso-ignore:padding;\n");
            htmlStr.Append("    color: windowtext;\n");
            htmlStr.Append("    font-size:14pt;\n");
            htmlStr.Append("    font-weight:400;\n");
            htmlStr.Append("    text-decoration:none;\n");
            htmlStr.Append("    font-family:'Times New Roman', serif;\n");
            htmlStr.Append("    mso-font-charset:0;\n");
            htmlStr.Append("    mso-number-format:General;\n");
            htmlStr.Append("    text-align:center;\n");
            htmlStr.Append("    vertical-align:middle;\n");
            htmlStr.Append("    border-top:none;\n");
            htmlStr.Append("    border-right:none;\n");
            htmlStr.Append("    border-bottom:none;\n");
            htmlStr.Append("    border-left:.5pt solid windowtext;\n");
            htmlStr.Append("    background: white;\n");
            htmlStr.Append("    mso-pattern:black none;\n");
            htmlStr.Append("    white-space:nowrap;\n");
            htmlStr.Append("}\n");
            htmlStr.Append(".xl10220842\n");
            htmlStr.Append("{\n");
            htmlStr.Append("    padding: 0px;\n");
            htmlStr.Append("    mso-ignore:padding;\n");
            htmlStr.Append("    color: windowtext;\n");
            htmlStr.Append("    font-size:14pt;\n");
            htmlStr.Append("    font-weight:400;\n");
            htmlStr.Append("    font-style:normal;\n");
            htmlStr.Append("    text-decoration:none;\n");
            htmlStr.Append("    font-family:'Times New Roman', serif;\n");
            htmlStr.Append("    mso-font-charset:0;\n");
            htmlStr.Append("    mso-number-format:General;\n");
            htmlStr.Append("    text-align:center;\n");
            htmlStr.Append("    vertical-align:middle;\n");
            htmlStr.Append("    background: white;\n");
            htmlStr.Append("    mso-pattern:black none;\n");
            htmlStr.Append("    white-space:nowrap;\n");
            htmlStr.Append("}\n");
            htmlStr.Append(".xl10820842\n");
            htmlStr.Append("{\n");
            htmlStr.Append("    padding: 0px;\n");
            htmlStr.Append("    mso-ignore:padding;\n");
            htmlStr.Append("    color: windowtext;\n");
            htmlStr.Append("    font-size:13.0pt;\n");
            htmlStr.Append("    font-weight:400;\n");
            htmlStr.Append("    font-style:normal;\n");
            htmlStr.Append("    text-decoration:none;\n");
            htmlStr.Append("    font-family:'Times New Roman', serif;\n");
            htmlStr.Append("    mso-font-charset:0;\n");
            htmlStr.Append("    mso-number-format:'_\\(* \\#\\,\\#\\#0_\\)\\;_\\(* \\\\(\\#\\,\\#\\#0\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';\n");
            htmlStr.Append("    text-align:center;\n");
            htmlStr.Append("    vertical-align:middle;\n");
            htmlStr.Append("    border-top:none;\n");
            htmlStr.Append("    border-right:.5pt solid windowtext;\n");
            htmlStr.Append("    border-bottom:none;\n");
            htmlStr.Append("    border-left:none;\n");
            htmlStr.Append("    background: white;\n");
            htmlStr.Append("    mso-pattern:black none;\n");
            htmlStr.Append("    white-space:nowrap;\n");
            htmlStr.Append("    mso-text-control:shrinktofit;\n");
            htmlStr.Append("}\n");
            htmlStr.Append(".xl11320842\n");
            htmlStr.Append("{\n");
            htmlStr.Append("    padding: 0px;\n");
            htmlStr.Append("    mso-ignore:padding;\n");
            htmlStr.Append("    color: windowtext;\n");
            htmlStr.Append("    font-size:13.0pt;\n");
            htmlStr.Append("    font-weight:400;\n");
            htmlStr.Append("    text-decoration:none;\n");
            htmlStr.Append("    font-family:'Times New Roman', serif;\n");
            htmlStr.Append("    mso-font-charset:0;\n");
            htmlStr.Append("    mso-number-format:General;\n");
            htmlStr.Append("    text-align:right;\n");
            htmlStr.Append("    vertical-align:middle;\n");
            htmlStr.Append("    background: white;\n");
            htmlStr.Append("    mso-pattern:black none;\n");
            htmlStr.Append("    white-space:nowrap;\n");
            htmlStr.Append("}\n");
            htmlStr.Append(".xl11120842\n");
            htmlStr.Append("{\n");
            htmlStr.Append("    padding: 0px;\n");
            htmlStr.Append("    mso-ignore:padding;\n");
            htmlStr.Append("    color: windowtext;\n");
            htmlStr.Append("    font-size:13.0pt;\n");
            htmlStr.Append("    font-weight:700;\n");
            htmlStr.Append("    font-style:normal;\n");
            htmlStr.Append("    text-decoration:none;\n");
            htmlStr.Append("    font-family:'Times New Roman', serif;\n");
            htmlStr.Append("    mso-font-charset:0;\n");
            htmlStr.Append("    mso-number-format:'_\\(* \\#\\,\\#\\#0\\.00_\\)\\;_\\(* \\\\(\\#\\,\\#\\#0\\.00\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';\n");
            htmlStr.Append("    text-align:center;\n");
            htmlStr.Append("    vertical-align:middle;\n");
            htmlStr.Append("    border-top:none;\n");
            htmlStr.Append("    border-right:.5pt solid windowtext;\n");
            htmlStr.Append("    border-bottom:none;\n");
            htmlStr.Append("    border-left:none;\n");
            htmlStr.Append("    background: white;\n");
            htmlStr.Append("    mso-pattern:black none;\n");
            htmlStr.Append("    white-space:nowrap;\n");
            htmlStr.Append("    mso-text-control:shrinktofit;\n");
            htmlStr.Append("}\n");
            htmlStr.Append(".xl11720842\n");
            htmlStr.Append("{\n");
            htmlStr.Append("    padding: 0px;\n");
            htmlStr.Append("    mso-ignore:padding;\n");
            htmlStr.Append("    color: windowtext;\n");
            htmlStr.Append("    font-size:13.0pt;\n");
            htmlStr.Append("    font-weight:400;\n");
            htmlStr.Append("    font-style:normal;\n");
            htmlStr.Append("    text-decoration:none;\n");
            htmlStr.Append("    font-family:'Times New Roman', serif;\n");
            htmlStr.Append("    mso-font-charset:0;\n");
            htmlStr.Append("    mso-number-format:'\\@';\n");
            htmlStr.Append("    text-align:general;\n");
            htmlStr.Append("    vertical-align:middle;\n");
            htmlStr.Append("    border-top:none;\n");
            htmlStr.Append("    border-right:.5pt solid windowtext;\n");
            htmlStr.Append("    border-bottom:none;\n");
            htmlStr.Append("    border-left:none;\n");
            htmlStr.Append("    background: white;\n");
            htmlStr.Append("    mso-pattern:black none;\n");
            htmlStr.Append("    white-space:nowrap;\n");
            htmlStr.Append("    mso-text-control:shrinktofit;\n");
            htmlStr.Append("}\n");
            htmlStr.Append(".xl10920842\n");
            htmlStr.Append("{\n");
            htmlStr.Append("    padding: 0px;\n");
            htmlStr.Append("    mso-ignore:padding;\n");
            htmlStr.Append("    color: windowtext;\n");
            htmlStr.Append("    font-size:13.0pt;\n");
            htmlStr.Append("    font-weight:400\n");
            htmlStr.Append("    font-style:normal;\n");
            htmlStr.Append("    text-decoration:none;\n");
            htmlStr.Append("    font-family:'Times New Roman', serif;\n");
            htmlStr.Append("    mso-font-charset:0;\n");
            htmlStr.Append("    mso-number-format:'_\\(* \\#\\,\\#\\#0\\.00_\\)\\;_\\(* \\\\(\\#\\,\\#\\#0\\.00\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';\n");
            htmlStr.Append("    text-align:center;\n");
            htmlStr.Append("    vertical-align:middle;\n");
            htmlStr.Append("    border-top:none;\n");
            htmlStr.Append("    border-right:none;\n");
            htmlStr.Append("    border-bottom:none;\n");
            htmlStr.Append("    border-left:.5pt solid windowtext;\n");
            htmlStr.Append("    background: white;\n");
            htmlStr.Append("    mso-pattern:black none;\n");
            htmlStr.Append("    white-space:nowrap;\n");
            htmlStr.Append("    mso-text-control:shrinktofit;\n");
            htmlStr.Append("}\n");
            htmlStr.Append(".xl12320842\n");
            htmlStr.Append("{\n");
            htmlStr.Append("    padding: 0px;\n");
            htmlStr.Append("    mso-ignore:padding;\n");
            htmlStr.Append("    color: windowtext;\n");
            htmlStr.Append("    font-size:13.0pt;\n");
            htmlStr.Append("    font-weight:400;\n");
            htmlStr.Append("    font-style:normal;\n");
            htmlStr.Append("    text-decoration:none;\n");
            htmlStr.Append("    font-family:'Times New Roman', serif;\n");
            htmlStr.Append("    mso-font-charset:0;\n");
            htmlStr.Append("    mso-number-format:'\\@';\n");
            htmlStr.Append("    text-align:right;\n");
            htmlStr.Append("    vertical-align:middle;\n");
            htmlStr.Append("    border-top:none;\n");
            htmlStr.Append("    border-right:.5pt solid windowtext;\n");
            htmlStr.Append("    border-bottom:none;\n");
            htmlStr.Append("    border-left:none;\n");
            htmlStr.Append("    background: white;\n");
            htmlStr.Append("    mso-pattern:black none;\n");
            htmlStr.Append("    white-space:normal;\n");
            htmlStr.Append("}\n");
            htmlStr.Append(".xl11620842\n");
            htmlStr.Append("{\n");
            htmlStr.Append("    padding: 0px;\n");
            htmlStr.Append("    mso-ignore:padding;\n");
            htmlStr.Append("    color: windowtext;\n");
            htmlStr.Append("    font-size:13.0pt;\n");
            htmlStr.Append("    font-weight:400;\n");
            htmlStr.Append("    font-style:normal;\n");
            htmlStr.Append("    text-decoration:none;\n");
            htmlStr.Append("    font-family:'Times New Roman', serif;\n");
            htmlStr.Append("    mso-font-charset:0;\n");
            htmlStr.Append("    mso-number-format:'\\@';\n");
            htmlStr.Append("    text-align:right;\n");
            htmlStr.Append("    vertical-align:middle;\n");
            htmlStr.Append("    border-top:none;\n");
            htmlStr.Append("    border-right:.5pt solid windowtext;\n");
            htmlStr.Append("    border-bottom:none;\n");
            htmlStr.Append("    border-left:none;\n");
            htmlStr.Append("    background: white;\n");
            htmlStr.Append("    mso-pattern:black none;\n");
            htmlStr.Append("    white-space:nowrap;\n");
            htmlStr.Append("    mso-text-control:shrinktofit;\n");
            htmlStr.Append("}\n");
            htmlStr.Append(".xl11520842\n");
            htmlStr.Append("{\n");
            htmlStr.Append("    padding: 0px;\n");
            htmlStr.Append("    mso-ignore:padding;\n");
            htmlStr.Append("    color: windowtext;\n");
            htmlStr.Append("    font-size:13.0pt;\n");
            htmlStr.Append("    font-weight:400;\n");
            htmlStr.Append("    font-style:normal;\n");
            htmlStr.Append("    text-decoration:none;\n");
            htmlStr.Append("    font-family:'Times New Roman', serif;\n");
            htmlStr.Append("    mso-font-charset:0;\n");
            htmlStr.Append("    mso-number-format:'\\@';\n");
            htmlStr.Append("    text-align:right;\n");
            htmlStr.Append("    vertical-align:middle;\n");
            htmlStr.Append("    border-top:none;\n");
            htmlStr.Append("    border-right:none;\n");
            htmlStr.Append("    border-bottom:none;\n");
            htmlStr.Append("    border-left:.5pt solid windowtext;\n");
            htmlStr.Append("    background: white;\n");
            htmlStr.Append("    mso-pattern:black none;\n");
            htmlStr.Append("    white-space:nowrap;\n");
            htmlStr.Append("    mso-text-control:shrinktofit;\n");
            htmlStr.Append("}\n");
            htmlStr.Append(".xl15520842\n");
            htmlStr.Append("{\n");
            htmlStr.Append("    padding: 0px;\n");
            htmlStr.Append("    mso-ignore:padding;\n");
            htmlStr.Append("    color: windowtext;\n");
            htmlStr.Append("    font-size:14pt;\n");
            htmlStr.Append("    font-weight:400;\n");
            htmlStr.Append("    font-style:normal;\n");
            htmlStr.Append("    text-decoration:none;\n");
            htmlStr.Append("    font-family:'Times New Roman', serif;\n");
            htmlStr.Append("    mso-font-charset:0;\n");
            htmlStr.Append("    mso-number-format:General;\n");
            htmlStr.Append("    text-align:center;\n");
            htmlStr.Append("    vertical-align:middle;\n");
            htmlStr.Append("    background: white;\n");
            htmlStr.Append("    mso-pattern:black none;\n");
            htmlStr.Append("    white-space:normal;\n");
            htmlStr.Append("}\n");
            htmlStr.Append(".xl8720842\n");
            htmlStr.Append("{\n");
            htmlStr.Append("    padding: 0px;\n");
            htmlStr.Append("    mso-ignore:padding;\n");
            htmlStr.Append("    color: windowtext;\n");
            htmlStr.Append("    font-size:14pt;\n");
            htmlStr.Append("    font-weight:400;\n");
            htmlStr.Append("    font-style:normal;\n");
            htmlStr.Append("    text-decoration:none;\n");
            htmlStr.Append("    font-family:'Times New Roman', serif;\n");
            htmlStr.Append("    mso-font-charset:0;\n");
            htmlStr.Append("    mso-number-format:General;\n");
            htmlStr.Append("    text-align:general;\n");
            htmlStr.Append("    vertical-align:middle;\n");
            htmlStr.Append("    background: white;\n");
            htmlStr.Append("    mso-pattern:black none;\n");
            htmlStr.Append("    white-space:nowrap;\n");
            htmlStr.Append("}\n");
            htmlStr.Append(".xl12420842\n");
            htmlStr.Append("{\n");
            htmlStr.Append("    padding: 0px;\n");
            htmlStr.Append("    mso-ignore:padding;\n");
            htmlStr.Append("    color: windowtext;\n");
            htmlStr.Append("    font-size:13.0pt;\n");
            htmlStr.Append("    font-weight:400;\n");
            htmlStr.Append("    font-style:normal;\n");
            htmlStr.Append("    text-decoration:none;\n");
            htmlStr.Append("    font-family:'Times New Roman', serif;\n");
            htmlStr.Append("    mso-font-charset:0;\n");
            htmlStr.Append("    mso-number-format:'\\@';\n");
            htmlStr.Append("    text-align:general;\n");
            htmlStr.Append("    vertical-align:middle;\n");
            htmlStr.Append("    border-top:none;\n");
            htmlStr.Append("    border-right:.5pt solid windowtext;\n");
            htmlStr.Append("    border-bottom:none;\n");
            htmlStr.Append("    border-left:none;\n");
            htmlStr.Append("    background: white;\n");
            htmlStr.Append("    mso-pattern:black none;\n");
            htmlStr.Append("    white-space:normal;\n");
            htmlStr.Append("}\n");

            htmlStr.Append(".xl14620842\n");
            htmlStr.Append("    {\n");
            htmlStr.Append("   padding: 0px;\n");
            htmlStr.Append("   mso-ignore:padding;\n");
            htmlStr.Append("   color: windowtext;\n");
            htmlStr.Append("   font-size:12.5pt;\n");
            htmlStr.Append("   font-weight:400;\n");
            htmlStr.Append("   font-style:italic;\n");
            htmlStr.Append("   text-decoration:none;\n");
            htmlStr.Append("   font-family:'Times New Roman', serif;\n");
            htmlStr.Append("   mso-font-charset:0;\n");
            htmlStr.Append("   mso-number-format:General;\n");
            htmlStr.Append("   text-align:center;\n");
            htmlStr.Append("   vertical-align:middle;\n");
            htmlStr.Append("   background: white;\n");
            htmlStr.Append("   mso-pattern:black none;\n");
            htmlStr.Append("   white-space:normal;\n");
            htmlStr.Append("   }\n");
            htmlStr.Append(".xl9120842 {                                                                                                                                                                                                  \n");
            htmlStr.Append("    padding: 0px;                                                                                                                                                                                             \n");
            htmlStr.Append("    mso-ignore: padding;                                                                                                                                                                                    \n");
            htmlStr.Append("    color: windowtext;                                                                                                                                                                                        \n");
            htmlStr.Append("    font-size: 10.0pt;                                                                                                                                                                                      \n");
            htmlStr.Append("    font-weight: 700;                                                                                                                                                                                       \n");
            htmlStr.Append("    font-style: normal;                                                                                                                                                                                     \n");
            htmlStr.Append("    text-decoration: none;                                                                                                                                                                                  \n");
            htmlStr.Append("    font-family: 'Times New Roman', serif;                                                                                                                                                                  \n");
            htmlStr.Append("    mso-font-charset: 0;                                                                                                                                                                                  \n");
            htmlStr.Append("    mso-number-format: General;                                                                                                                                                                           \n");
            htmlStr.Append("    text-align: general;                                                                                                                                                                                    \n");
            htmlStr.Append("    vertical-align: middle;                                                                                                                                                                                 \n");
            htmlStr.Append("    background: white;                                                                                                                                                                                        \n");
            htmlStr.Append("    mso-pattern: black none;                                                                                                                                                                                \n");
            htmlStr.Append("    white-space: nowrap;                                                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                                                             \n");
            htmlStr.Append(".xl14120842 {                                                                                                                                                                                                 \n");
            htmlStr.Append("    padding: 0px;                                                                                                                                                                                             \n");
            htmlStr.Append("    mso-ignore: padding;                                                                                                                                                                                    \n");
            htmlStr.Append("    color: windowtext;                                                                                                                                                                                        \n");
            htmlStr.Append("    font-size: 10.0pt;                                                                                                                                                                                      \n");
            htmlStr.Append("    font-weight: 700;                                                                                                                                                                                       \n");
            htmlStr.Append("    font-style: normal;                                                                                                                                                                                     \n");
            htmlStr.Append("    text-decoration: none;                                                                                                                                                                                  \n");
            htmlStr.Append("    font-family: 'Times New Roman', serif;                                                                                                                                                                  \n");
            htmlStr.Append("    mso-font-charset: 0;                                                                                                                                                                                  \n");
            htmlStr.Append("    mso-number-format: General;                                                                                                                                                                           \n");
            htmlStr.Append("    text-align: general;                                                                                                                                                                                    \n");
            htmlStr.Append("    vertical-align: middle;                                                                                                                                                                                 \n");
            htmlStr.Append("    border-top: none;                                                                                                                                                                                       \n");
            htmlStr.Append("    border-right: none;                                                                                                                                                                                     \n");
            htmlStr.Append("    border-bottom: none;                                                                                                                                                                                    \n");
            htmlStr.Append("    border-left: .5pt solid windowtext;                                                                                                                                                                     \n");
            htmlStr.Append("    background: white;                                                                                                                                                                                        \n");
            htmlStr.Append("    mso-pattern: black none;                                                                                                                                                                                \n");
            htmlStr.Append("    white-space: nowrap;                                                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                                                             \n");
            htmlStr.Append(".xl18320842 {                                                                                                                                                                                                 \n");
            htmlStr.Append("    padding: 0px;                                                                                                                                                                                             \n");
            htmlStr.Append("    mso-ignore: padding;                                                                                                                                                                                    \n");
            htmlStr.Append("    color: windowtext;                                                                                                                                                                                        \n");
            htmlStr.Append("    font-size: 11.0pt;                                                                                                                                                                                      \n");
            htmlStr.Append("    font-weight: 700;                                                                                                                                                                                       \n");
            htmlStr.Append("    font-style: normal;                                                                                                                                                                                     \n");
            htmlStr.Append("    text-decoration: none;                                                                                                                                                                                  \n");
            htmlStr.Append("    font-family: 'Times New Roman', serif;                                                                                                                                                                  \n");
            htmlStr.Append("    mso-font-charset: 0;                                                                                                                                                                                  \n");
            htmlStr.Append("    mso-number-format: General;                                                                                                                                                                           \n");
            htmlStr.Append("    text-align: left;                                                                                                                                                                                       \n");
            htmlStr.Append("    vertical-align: middle;                                                                                                                                                                                 \n");
            htmlStr.Append("    mso-background-source: auto;                                                                                                                                                                          \n");
            htmlStr.Append("    mso-pattern: auto;                                                                                                                                                                                      \n");
            htmlStr.Append("    white-space: normal;                                                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                                                             \n");
            htmlStr.Append(".xl18420842 {                                                                                                                                                                                                 \n");
            htmlStr.Append("    padding: 0px;                                                                                                                                                                                             \n");
            htmlStr.Append("    mso-ignore: padding;                                                                                                                                                                                    \n");
            htmlStr.Append("    color: windowtext;                                                                                                                                                                                        \n");
            htmlStr.Append("    font-size: 12.0pt;                                                                                                                                                                                      \n");
            htmlStr.Append("    font-weight: 700;                                                                                                                                                                                       \n");
            htmlStr.Append("    font-style: normal;                                                                                                                                                                                     \n");
            htmlStr.Append("    text-decoration: none;                                                                                                                                                                                  \n");
            htmlStr.Append("    font-family: 'Times New Roman', serif;                                                                                                                                                                  \n");
            htmlStr.Append("    mso-font-charset: 0;                                                                                                                                                                                  \n");
            htmlStr.Append("    mso-number-format: General;                                                                                                                                                                           \n");
            htmlStr.Append("    text-align: center;                                                                                                                                                                                     \n");
            htmlStr.Append("    vertical-align: middle;                                                                                                                                                                                 \n");
            htmlStr.Append("    mso-background-source: auto;                                                                                                                                                                          \n");
            htmlStr.Append("    mso-pattern: auto;                                                                                                                                                                                      \n");
            htmlStr.Append("    white-space: normal;                                                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                                                             \n");
            htmlStr.Append(".xl17620842 {                                                                                                                                                                                                 \n");
            htmlStr.Append("    padding: 0px;                                                                                                                                                                                             \n");
            htmlStr.Append("    mso-ignore: padding;                                                                                                                                                                                    \n");
            htmlStr.Append("    color: windowtext;                                                                                                                                                                                        \n");
            htmlStr.Append("    font-size: 11.0pt;                                                                                                                                                                                      \n");
            htmlStr.Append("    font-weight: 700;                                                                                                                                                                                       \n");
            htmlStr.Append("    font-style: normal;                                                                                                                                                                                     \n");
            htmlStr.Append("    text-decoration: none;                                                                                                                                                                                  \n");
            htmlStr.Append("    font-family: 'Times New Roman', serif;                                                                                                                                                                  \n");
            htmlStr.Append("    mso-font-charset: 0;                                                                                                                                                                                  \n");
            htmlStr.Append("    mso-number-format: General;                                                                                                                                                                           \n");
            htmlStr.Append("    text-align: center;                                                                                                                                                                                     \n");
            htmlStr.Append("    vertical-align: middle;                                                                                                                                                                                 \n");
            htmlStr.Append("    background: white;                                                                                                                                                                                        \n");
            htmlStr.Append("    mso-pattern: black none;                                                                                                                                                                                \n");
            htmlStr.Append("    white-space: normal;                                                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                                                             \n");
            htmlStr.Append(".xl14020842 {                                                                                                                                                                                                 \n");
            htmlStr.Append("    padding: 0px;                                                                                                                                                                                             \n");
            htmlStr.Append("    mso-ignore: padding;                                                                                                                                                                                    \n");
            htmlStr.Append("    color: windowtext;                                                                                                                                                                                        \n");
            htmlStr.Append("    font-size: 12.0pt;                                                                                                                                                                                      \n");
            htmlStr.Append("    font-weight: 700;                                                                                                                                                                                       \n");
            htmlStr.Append("    font-style: normal;                                                                                                                                                                                     \n");
            htmlStr.Append("    text-decoration: none;                                                                                                                                                                                  \n");
            htmlStr.Append("    font-family: 'Times New Roman', serif;                                                                                                                                                                  \n");
            htmlStr.Append("    mso-font-charset: 0;                                                                                                                                                                                  \n");
            htmlStr.Append("    mso-number-format: General;                                                                                                                                                                           \n");
            htmlStr.Append("    text-align: general;                                                                                                                                                                                    \n");
            htmlStr.Append("    vertical-align: middle;                                                                                                                                                                                 \n");
            htmlStr.Append("    border-top: none;                                                                                                                                                                                       \n");
            htmlStr.Append("    border-right: .5pt solid windowtext;                                                                                                                                                                    \n");
            htmlStr.Append("    border-bottom: none;                                                                                                                                                                                    \n");
            htmlStr.Append("    border-left: none;                                                                                                                                                                                      \n");
            htmlStr.Append("    background: white;                                                                                                                                                                                        \n");
            htmlStr.Append("    mso-pattern: black none;                                                                                                                                                                                \n");
            htmlStr.Append("    white-space: normal;                                                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                                                             \n");
            htmlStr.Append(".xl8120842 {                                                                                                                                                                                                  \n");
            htmlStr.Append("    padding: 0px;                                                                                                                                                                                             \n");
            htmlStr.Append("    mso-ignore: padding;                                                                                                                                                                                    \n");
            htmlStr.Append("    color: windowtext;                                                                                                                                                                                        \n");
            htmlStr.Append("    font-size: 12.0pt;                                                                                                                                                                                      \n");
            htmlStr.Append("    font-weight: 400;                                                                                                                                                                                       \n");
            htmlStr.Append("    font-style: normal;                                                                                                                                                                                     \n");
            htmlStr.Append("    text-decoration: none;                                                                                                                                                                                  \n");
            htmlStr.Append("    font-family: 'Times New Roman', serif;                                                                                                                                                                  \n");
            htmlStr.Append("    mso-font-charset: 0;                                                                                                                                                                                  \n");
            htmlStr.Append("    mso-number-format: General;                                                                                                                                                                           \n");
            htmlStr.Append("    text-align: general;                                                                                                                                                                                    \n");
            htmlStr.Append("    vertical-align: middle;                                                                                                                                                                                 \n");
            htmlStr.Append("    background: white;                                                                                                                                                                                        \n");
            htmlStr.Append("    mso-pattern: black none;                                                                                                                                                                                \n");
            htmlStr.Append("    white-space: nowrap;                                                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                                                             \n");
            htmlStr.Append(".xl14220842 {                                                                                                                                                                                                 \n");
            htmlStr.Append("    padding: 0px;                                                                                                                                                                                             \n");
            htmlStr.Append("    mso-ignore: padding;                                                                                                                                                                                    \n");
            htmlStr.Append("    color: windowtext;                                                                                                                                                                                        \n");
            htmlStr.Append("    font-size: 11.0pt;                                                                                                                                                                                      \n");
            htmlStr.Append("    font-weight: 400;                                                                                                                                                                                       \n");
            htmlStr.Append("    font-style: normal;                                                                                                                                                                                     \n");
            htmlStr.Append("    text-decoration: none;                                                                                                                                                                                  \n");
            htmlStr.Append("    font-family: 'Times New Roman', serif;                                                                                                                                                                  \n");
            htmlStr.Append("    mso-font-charset: 0;                                                                                                                                                                                  \n");
            htmlStr.Append("    mso-number-format: General;                                                                                                                                                                           \n");
            htmlStr.Append("    text-align: general;                                                                                                                                                                                    \n");
            htmlStr.Append("    vertical-align: middle;                                                                                                                                                                                 \n");
            htmlStr.Append("    border-top: none;                                                                                                                                                                                       \n");
            htmlStr.Append("    border-right: none;                                                                                                                                                                                     \n");
            htmlStr.Append("    border-bottom: none;                                                                                                                                                                                    \n");
            htmlStr.Append("    border-left: .5pt solid windowtext;                                                                                                                                                                     \n");
            htmlStr.Append("    background: white;                                                                                                                                                                                        \n");
            htmlStr.Append("    mso-pattern: black none;                                                                                                                                                                                \n");
            htmlStr.Append("    white-space: nowrap;                                                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                                                             \n");
            htmlStr.Append(".xl18520842 {                                                                                                                                                                                                 \n");
            htmlStr.Append("    padding: 0px;                                                                                                                                                                                             \n");
            htmlStr.Append("    mso-ignore: padding;                                                                                                                                                                                    \n");
            htmlStr.Append("    color: windowtext;                                                                                                                                                                                        \n");
            htmlStr.Append("    font-size: 11.0pt;                                                                                                                                                                                      \n");
            htmlStr.Append("    font-weight: 400;                                                                                                                                                                                       \n");
            htmlStr.Append("    font-style: normal;                                                                                                                                                                                     \n");
            htmlStr.Append("    text-decoration: none;                                                                                                                                                                                  \n");
            htmlStr.Append("    font-family: 'Times New Roman', serif;                                                                                                                                                                  \n");
            htmlStr.Append("    mso-font-charset: 0;                                                                                                                                                                                  \n");
            htmlStr.Append("    mso-number-format: General;                                                                                                                                                                           \n");
            htmlStr.Append("    text-align: center;                                                                                                                                                                                     \n");
            htmlStr.Append("    vertical-align: middle;                                                                                                                                                                                 \n");
            htmlStr.Append("    mso-background-source: auto;                                                                                                                                                                          \n");
            htmlStr.Append("    mso-pattern: auto;                                                                                                                                                                                      \n");
            htmlStr.Append("    white-space: normal;                                                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                                                             \n");
            htmlStr.Append(".xl14320842 {                                                                                                                                                                                                 \n");
            htmlStr.Append("    padding: 0px;                                                                                                                                                                                             \n");
            htmlStr.Append("    mso-ignore: padding;                                                                                                                                                                                    \n");
            htmlStr.Append("    color: windowtext;                                                                                                                                                                                        \n");
            htmlStr.Append("    font-size: 11.0pt;                                                                                                                                                                                      \n");
            htmlStr.Append("    font-weight: 400;                                                                                                                                                                                       \n");
            htmlStr.Append("    font-style: normal;                                                                                                                                                                                     \n");
            htmlStr.Append("    text-decoration: none;                                                                                                                                                                                  \n");
            htmlStr.Append("    font-family: 'Times New Roman', serif;                                                                                                                                                                  \n");
            htmlStr.Append("    mso-font-charset: 0;                                                                                                                                                                                  \n");
            htmlStr.Append("    mso-number-format: General;                                                                                                                                                                           \n");
            htmlStr.Append("    text-align: center;                                                                                                                                                                                     \n");
            htmlStr.Append("    vertical-align: middle;                                                                                                                                                                                 \n");
            htmlStr.Append("    background: white;                                                                                                                                                                                        \n");
            htmlStr.Append("    mso-pattern: black none;                                                                                                                                                                                \n");
            htmlStr.Append("    white-space: normal;                                                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                                                             \n");
            htmlStr.Append(".xl14420842 {                                                                                                                                                                                                 \n");
            htmlStr.Append("    padding: 0px;                                                                                                                                                                                             \n");
            htmlStr.Append("    mso-ignore: padding;                                                                                                                                                                                    \n");
            htmlStr.Append("    color: windowtext;                                                                                                                                                                                        \n");
            htmlStr.Append("    font-size: 11.0pt;                                                                                                                                                                                      \n");
            htmlStr.Append("    font-weight: 400;                                                                                                                                                                                       \n");
            htmlStr.Append("    font-style: italic;                                                                                                                                                                                     \n");
            htmlStr.Append("    text-decoration: none;                                                                                                                                                                                  \n");
            htmlStr.Append("    font-family: 'Times New Roman', serif;                                                                                                                                                                  \n");
            htmlStr.Append("    mso-font-charset: 0;                                                                                                                                                                                  \n");
            htmlStr.Append("    mso-number-format: General;                                                                                                                                                                           \n");
            htmlStr.Append("    text-align: general;                                                                                                                                                                                    \n");
            htmlStr.Append("    vertical-align: middle;                                                                                                                                                                                 \n");
            htmlStr.Append("    border-top: none;                                                                                                                                                                                       \n");
            htmlStr.Append("    border-right: .5pt solid windowtext;                                                                                                                                                                    \n");
            htmlStr.Append("    border-bottom: none;                                                                                                                                                                                    \n");
            htmlStr.Append("    border-left: none;                                                                                                                                                                                      \n");
            htmlStr.Append("    background: white;                                                                                                                                                                                        \n");
            htmlStr.Append("    mso-pattern: black none;                                                                                                                                                                                \n");
            htmlStr.Append("    white-space: normal;                                                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                                                             \n");
            htmlStr.Append(".xl14820842 {                                                                                                                                                                                                 \n");
            htmlStr.Append("    padding: 0px;                                                                                                                                                                                             \n");
            htmlStr.Append("    mso-ignore: padding;                                                                                                                                                                                    \n");
            htmlStr.Append("    color: windowtext;                                                                                                                                                                                        \n");
            htmlStr.Append("    font-size: 10.0pt;                                                                                                                                                                                      \n");
            htmlStr.Append("    font-weight: 400;                                                                                                                                                                                       \n");
            htmlStr.Append("    font-style: italic;                                                                                                                                                                                     \n");
            htmlStr.Append("    text-decoration: none;                                                                                                                                                                                  \n");
            htmlStr.Append("    font-family: 'Times New Roman', serif;                                                                                                                                                                  \n");
            htmlStr.Append("    mso-font-charset: 0;                                                                                                                                                                                  \n");
            htmlStr.Append("    mso-number-format: General;                                                                                                                                                                           \n");
            htmlStr.Append("    text-align: general;                                                                                                                                                                                    \n");
            htmlStr.Append("    vertical-align: middle;                                                                                                                                                                                 \n");
            htmlStr.Append("    background: white;                                                                                                                                                                                        \n");
            htmlStr.Append("    mso-pattern: black none;                                                                                                                                                                                \n");
            htmlStr.Append("    white-space: nowrap;                                                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                                                             \n");
            htmlStr.Append(".xl14520842 {                                                                                                                                                                                                 \n");
            htmlStr.Append("    padding: 0px;                                                                                                                                                                                             \n");
            htmlStr.Append("    mso-ignore: padding;                                                                                                                                                                                    \n");
            htmlStr.Append("    color: windowtext;                                                                                                                                                                                        \n");
            htmlStr.Append("    font-size: 10.0pt;                                                                                                                                                                                      \n");
            htmlStr.Append("    font-weight: 400;                                                                                                                                                                                       \n");
            htmlStr.Append("    font-style: italic;                                                                                                                                                                                     \n");
            htmlStr.Append("    text-decoration: none;                                                                                                                                                                                  \n");
            htmlStr.Append("    font-family: 'Times New Roman', serif;                                                                                                                                                                  \n");
            htmlStr.Append("    mso-font-charset: 0;                                                                                                                                                                                  \n");
            htmlStr.Append("    mso-number-format: General;                                                                                                                                                                           \n");
            htmlStr.Append("    text-align: general;                                                                                                                                                                                    \n");
            htmlStr.Append("    vertical-align: middle;                                                                                                                                                                                 \n");
            htmlStr.Append("    border-top: none;                                                                                                                                                                                       \n");
            htmlStr.Append("    border-right: none;                                                                                                                                                                                     \n");
            htmlStr.Append("    border-bottom: none;                                                                                                                                                                                    \n");
            htmlStr.Append("    border-left: .5pt solid windowtext;                                                                                                                                                                     \n");
            htmlStr.Append("    background: white;                                                                                                                                                                                        \n");
            htmlStr.Append("    mso-pattern: black none;                                                                                                                                                                                \n");
            htmlStr.Append("    white-space: nowrap;                                                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                                                             \n");
            htmlStr.Append(".xl18620842 {                                                                                                                                                                                                 \n");
            htmlStr.Append("    padding: 0px;                                                                                                                                                                                             \n");
            htmlStr.Append("    mso-ignore: padding;                                                                                                                                                                                    \n");
            htmlStr.Append("    color: windowtext;                                                                                                                                                                                        \n");
            htmlStr.Append("    font-size: 12.0pt;                                                                                                                                                                                      \n");
            htmlStr.Append("    font-weight: 400;                                                                                                                                                                                       \n");
            htmlStr.Append("    font-style: italic;                                                                                                                                                                                     \n");
            htmlStr.Append("    text-decoration: none;                                                                                                                                                                                  \n");
            htmlStr.Append("    font-family: 'Times New Roman', serif;                                                                                                                                                                  \n");
            htmlStr.Append("    mso-font-charset: 0;                                                                                                                                                                                  \n");
            htmlStr.Append("    mso-number-format: General;                                                                                                                                                                           \n");
            htmlStr.Append("    text-align: center;                                                                                                                                                                                     \n");
            htmlStr.Append("    vertical-align: middle;                                                                                                                                                                                 \n");
            htmlStr.Append("    mso-background-source: auto;                                                                                                                                                                          \n");
            htmlStr.Append("    mso-pattern: auto;                                                                                                                                                                                      \n");
            htmlStr.Append("    white-space: normal;                                                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                                                             \n");
            htmlStr.Append(".xl14720842 {                                                                                                                                                                                                 \n");
            htmlStr.Append("    padding: 0px;                                                                                                                                                                                             \n");
            htmlStr.Append("    mso-ignore: padding;                                                                                                                                                                                    \n");
            htmlStr.Append("    color: windowtext;                                                                                                                                                                                        \n");
            htmlStr.Append("    font-size: 12.0pt;                                                                                                                                                                                      \n");
            htmlStr.Append("    font-weight: 400;                                                                                                                                                                                       \n");
            htmlStr.Append("    font-style: italic;                                                                                                                                                                                     \n");
            htmlStr.Append("    text-decoration: none;                                                                                                                                                                                  \n");
            htmlStr.Append("    font-family: 'Times New Roman', serif;                                                                                                                                                                  \n");
            htmlStr.Append("    mso-font-charset: 0;                                                                                                                                                                                  \n");
            htmlStr.Append("    mso-number-format: General;                                                                                                                                                                           \n");
            htmlStr.Append("    text-align: general;                                                                                                                                                                                    \n");
            htmlStr.Append("    vertical-align: middle;                                                                                                                                                                                 \n");
            htmlStr.Append("    border-top: none;                                                                                                                                                                                       \n");
            htmlStr.Append("    border-right: .5pt solid windowtext;                                                                                                                                                                    \n");
            htmlStr.Append("    border-bottom: none;                                                                                                                                                                                    \n");
            htmlStr.Append("    border-left: none;                                                                                                                                                                                      \n");
            htmlStr.Append("    background: white;                                                                                                                                                                                        \n");
            htmlStr.Append("    mso-pattern: black none;                                                                                                                                                                                \n");
            htmlStr.Append("    white-space: normal;                                                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                                                             \n");
            htmlStr.Append("                                                                                                                                                                                                              \n");
            htmlStr.Append("-->                                                                                                                                                \n");
            htmlStr.Append("</style>                                                                                                                                           \n");
            htmlStr.Append("</head>                                                                                                                                            \n");
            htmlStr.Append("                                                                                                                                                   \n");
            htmlStr.Append("<body>                                                                                                                                             \n");
            htmlStr.Append("<!--[if !excel]>&nbsp;&nbsp;<![endif]-->                                                                                                           \n");
            htmlStr.Append("<!--The following information was generated by Microsoft Excel's Publish as Web                                                                    \n");
            htmlStr.Append("Page wizard.-->                                                                                                                                    \n");
            htmlStr.Append("<!--If the same item is republished from Excel, all information between the DIV                                                                    \n");
            htmlStr.Append("tags will be replaced.-->                                                                                                                          \n");
            htmlStr.Append("<!----------------------------->                                                                                                                   \n");
            htmlStr.Append("<!--START OF OUTPUT FROM EXCEL PUBLISH AS WEB PAGE WIZARD -->                                                                                      \n");
            htmlStr.Append("<!----------------------------->                                                                                                                   \n");
            htmlStr.Append("                                                                                                                                                   \n");

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

            double v_totalHeightLastPage = 323.5;// 258.5;

            double v_totalHeightPage = 580;//   540;

            for (int k = 0; k < v_countNumberOfPages; k++)
            {
                v_totalHeightPage = 580;// 540;

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

                htmlStr.Append("<div id='hóa &#273;&#417;n m&#7851;u suheung_15939' align=center                                                                                   \n");
                htmlStr.Append("x:publishsource='Excel'>                                                                                                                           \n");
                htmlStr.Append("    \n");
                htmlStr.Append("<table border=1 cellpadding=0 cellspacing=0 width=709 class=xl7315939                                                                              \n");
                htmlStr.Append(" style='border-collapse:collapse;table-layout:fixed;width:660pt'>                                                                                  \n");
                htmlStr.Append(" <col class=xl7315939 width=6 style='mso-width-source:userset;mso-width-alt:                                                                       \n");
                htmlStr.Append(" 199;width:5pt'>                                                                                                                                   \n");
                htmlStr.Append(" <col class=xl7315939 width=33 style='mso-width-source:userset;mso-width-alt:                                                                      \n");
                htmlStr.Append(" 1166;width:37.5pt'>                                                                                                                                 \n");
                htmlStr.Append(" <col class=xl7315939 width=127 style='mso-width-source:userset;mso-width-alt:                                                                     \n");
                htmlStr.Append(" 4522;width:81.31.25pt'>                                                                                                                                 \n");
                htmlStr.Append(" <col class=xl7315939 width=62 style='mso-width-source:userset;mso-width-alt:                                                                      \n");
                htmlStr.Append(" 2218;width:71.31.25pt'>                                                                                                                                 \n");
                htmlStr.Append(" <col class=xl7315939 width=41 style='mso-width-source:userset;mso-width-alt:                                                                      \n");
                htmlStr.Append(" 1450;width:51pt'>                                                                                                                                 \n");
                htmlStr.Append(" <col class=xl7315939 width=34 style='mso-width-source:userset;mso-width-alt:                                                                      \n");
                htmlStr.Append(" 1194;width:31.25pt'>                                                                                                                                 \n");
                htmlStr.Append(" <col class=xl7315939 width=77 style='mso-width-source:userset;mso-width-alt:                                                                      \n");
                htmlStr.Append(" 2730;width:82.5‬pt'>                                                                                                                                 \n");
                htmlStr.Append(" <col class=xl7315939 width=78 style='mso-width-source:userset;mso-width-alt:                                                                      \n");
                htmlStr.Append(" 2759;width:72.5‬pt'>                                                                                                                                 \n");
                htmlStr.Append(" <col class=xl7315939 width=6 style='mso-width-source:userset;mso-width-alt:                                                                       \n");
                htmlStr.Append(" 199;width:5pt'>                                                                                                                                   \n");
                htmlStr.Append(" <col class=xl7315939 width=49 style='mso-width-source:userset;mso-width-alt:                                                                      \n");
                htmlStr.Append(" 1735;width:46.25‬pt'>                                                                                                                                 \n");
                htmlStr.Append(" <col class=xl7315939 width=48 style='mso-width-source:userset;mso-width-alt:                                                                      \n");
                htmlStr.Append(" 1706;width:45pt'>                                                                                                                                 \n");
                htmlStr.Append(" <col class=xl7315939 width=6 style='mso-width-source:userset;mso-width-alt:                                                                       \n");
                htmlStr.Append(" 199;width:5pt'>                                                                                                                                   \n");
                htmlStr.Append(" <col class=xl7315939 width=45 style='mso-width-source:userset;mso-width-alt:                                                                      \n");
                htmlStr.Append(" 1592;width:53.75pt'>                                                                                                                                 \n");
                htmlStr.Append(" <col class=xl7315939 width=85 style='mso-width-source:userset;mso-width-alt:                                                                      \n");
                htmlStr.Append(" 3015;width:51.25pt'>                                                                                                                                 \n");
                htmlStr.Append(" <col class=xl7315939 width=12 style='mso-width-source:userset;mso-width-alt:                                                                      \n");
                htmlStr.Append(" 426;width:18.75ptpt'>                                                                                                                                      \n");
                htmlStr.Append(" <tr height=6 style='mso-height-source:userset;height:5.1pt'>                                                                                      \n");
                htmlStr.Append("  <td height=6 class=xl7415939 width=6 style='height:5.1pt;width:5pt;font-size:0pt'>&nbsp;</td>                                                                  \n");
                htmlStr.Append("  <td class=xl7515939 width=33 style='width:31.25pt;font-size:0pt'>&nbsp;</td>                                                                                      \n");
                htmlStr.Append("  <td class=xl7515939 width=127 style='width:95pt;font-size:0pt'>&nbsp;</td>                                                                                     \n");
                htmlStr.Append("  <td class=xl7515939 width=62 style='width:47pt;font-size:0pt'>&nbsp;</td>                                                                                      \n");
                htmlStr.Append("  <td class=xl7515939 width=41 style='width:31pt;font-size:0pt'>&nbsp;</td>                                                                                      \n");
                htmlStr.Append("  <td class=xl7515939 width=34 style='width:31.25pt;font-size:0pt'>&nbsp;</td>                                                                                      \n");
                htmlStr.Append("  <td class=xl7515939 width=77 style='width:72.5‬pt;font-size:0pt'>&nbsp;</td>                                                                                      \n");
                htmlStr.Append("  <td class=xl7515939 width=78 style='width:72.5‬pt;font-size:0pt'>&nbsp;</td>                                                                                      \n");
                htmlStr.Append("  <td class=xl7515939 width=6 style='width:5pt;font-size:0pt'>&nbsp;</td>                                                                                        \n");
                htmlStr.Append("  <td class=xl7515939 width=49 style='width:46.25‬pt;font-size:0pt'>&nbsp;</td>                                                                                      \n");
                htmlStr.Append("  <td class=xl7515939 width=48 style='width:45pt;font-size:0pt'>&nbsp;</td>                                                                                      \n");
                htmlStr.Append("  <td class=xl7515939 width=6 style='width:5pt;font-size:0pt'>&nbsp;</td>                                                                                        \n");
                htmlStr.Append("  <td class=xl7515939 width=45 style='width:43.75pt;font-size:0pt'>&nbsp;</td>                                                                                      \n");
                htmlStr.Append("  <td class=xl7515939 width=85 style='width:81.31.25pt;font-size:0pt'>&nbsp;</td>                                                                                      \n");
                htmlStr.Append("  <td class=xl7615939 width=12 style='width:9pt;font-size:0pt'>&nbsp;</td>                                                                                       \n");
                htmlStr.Append(" </tr>                                                                                                                                             \n");
                htmlStr.Append(" <tr height=30 style='height:28.5‬pt'>                                                                                                              \n");
                htmlStr.Append("  <td height=30 class=xl8415939 style='height:28.5‬pt'>&nbsp;</td>                                                                                  \n");
                htmlStr.Append("  <td class=xl7715939 style='border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                                                  \n");

                if (Double.Parse(dt.Rows[0]["INVOICE_DATE_110"].ToString()) <= 20230512)
                {
                    htmlStr.Append("  <td align=left style='border-top:none;border-bottom:none;border-left:none;border-right:none' valign=top><![if !vml]><span style='mso-ignore:vglayout;                                                                          \n");
                    htmlStr.Append("  position:absolute;z-index:2;margin-left:7px;margin-top:3px;width:90px;                                                                           \n");
                    htmlStr.Append("  height:101px'><img width=90 height=101                                                                                                           \n");
                    htmlStr.Append("  src='D:/webproject/e-invoice-ws/02.Web/EInvoice/img/Suheung_1.png'                                                                 \n");
                    htmlStr.Append("  v:shapes='Picture_x0020_6'></span><![endif]><span style='mso-ignore:vglayout2'>                                                                  \n");
                    htmlStr.Append("  <table cellpadding=0 cellspacing=0>                                                                                                              \n");
                    htmlStr.Append("   <tr>                                                                                                                                            \n");
                    htmlStr.Append("    <td height=30 class=xl7315939 width=127 style='height:28.5‬pt;width:95pt;none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                           \n");
                    htmlStr.Append("   </tr>                                                                                                                                           \n");
                    htmlStr.Append("  </table>                                                                                                                                         \n");
                    htmlStr.Append("  </span></td>                                                                                                                                     \n");

                }
                else
                {
                    htmlStr.Append("  <td align=left style='border-top:none;border-bottom:none;border-left:none;border-right:none' valign=top><![if !vml]><span style='mso-ignore:vglayout;                                                                          \n");
                    htmlStr.Append("  position:absolute;z-index:2;margin-left:7px;margin-top:3px;width:90px;                                                                           \n");
                    htmlStr.Append("  height:101px'><img width=90 height=101                                                                                                           \n");
                    htmlStr.Append("  src='D:/webproject/e-invoice-ws/02.Web/EInvoice/img/Suheung_1_20230512.jpg'                                                                 \n");
                    htmlStr.Append("  v:shapes='Picture_x0020_6'></span><![endif]><span style='mso-ignore:vglayout2'>                                                                  \n");
                    htmlStr.Append("  <table cellpadding=0 cellspacing=0>                                                                                                              \n");
                    htmlStr.Append("   <tr>                                                                                                                                            \n");
                    htmlStr.Append("    <td height=30 class=xl7315939 width=127 style='height:28.5‬pt;width:95pt;none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                           \n");
                    htmlStr.Append("   </tr>                                                                                                                                           \n");
                    htmlStr.Append("  </table>                                                                                                                                         \n");
                    htmlStr.Append("  </span></td>                                                                                                                                     \n");

                }
                htmlStr.Append("  <td colspan=7 style='border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl15415939>HÓA ĐƠN GIÁ TRỊ GIA TĂNG</td>                                                                \n");
                htmlStr.Append("  <td style='border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl7815939>&nbsp;</td>                                                                                                                  \n");
                htmlStr.Append("  <td style='border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl7815939>&nbsp;</td>                                                                                                                  \n");
                htmlStr.Append("  <td  style='border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl7715939>&nbsp;</td>                                                                                                                  \n");
                htmlStr.Append("  <td style='border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl7715939>&nbsp;</td>                                                                                                                  \n");
                htmlStr.Append("  <td style='border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl7915939>&nbsp;</td>                                                                                                                  \n");
                htmlStr.Append(" </tr>                                                                                                                                             \n");
                htmlStr.Append(" <tr height=27 style='height:25.625pt'>                                                                                                              \n");
                htmlStr.Append("  <td height=27 class=xl8415939 style='height:25.625pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                  \n");
                htmlStr.Append("  <td style='border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl7315939>&nbsp;</td>                                                                                                                  \n");
                htmlStr.Append("  <td style='border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl7315939>&nbsp;</td>                                                                                                                  \n");
                htmlStr.Append("  <td style='border-top:none;border-bottom:none;border-left:none;border-right:none' colspan=7 class=xl15515939>VAT INVOICE</td>                                                                                                  \n");
                htmlStr.Append("  <td class=xl8015939 colspan=5 style='border-top:none;border-bottom:none;border-left:none;border-right:none'><font class='font1015939'></font><font class='font815939'></font><font class='font715939'></font></td>                     \n");
                htmlStr.Append(" </tr>                                                                                                                                             \n");
                htmlStr.Append(" <tr height=24 style='mso-height-source:userset;height:22.5pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>                                                                                    \n");
                htmlStr.Append("  <td height=24 class=xl8415939 style='height:22.5pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                  \n");
                htmlStr.Append("  <td style='border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl7315939>&nbsp;</td>                                                                                                                  \n");
                htmlStr.Append("  <td style='border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl7315939>&nbsp;</td>                                                                                                                  \n");
                htmlStr.Append("  <td style='border-top:none;border-bottom:none;border-left:none;border-right:none' colspan=7 class=xl18915939>(HÓA &#272;&#416;N CHUY&#7874;N &#272;&#7892;I T&#7914; HÓA &#272;&#416;N &#272;I&#7878;N T&#7916;)<span                                                                                                              \n");
                htmlStr.Append("  style='mso-spacerun:yes'> </span></td>                                                                                                           \n");
                htmlStr.Append("  <td style='border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl8015939 colspan=4>Ký hiệu (<font class='font1015939'>Serial</font><font                                                        \n");
                htmlStr.Append("  class='font815939'>):</font><font class='font715939'> " + dt.Rows[0]["templateCode"] + "" + dt.Rows[0]["InvoiceSerialNo"] + "</font></td>                                                                         \n");
                htmlStr.Append("  <td style='border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl8115939>&nbsp;</td>                                                                                                                  \n");
                htmlStr.Append(" </tr>                                                                                                                                             \n");
                htmlStr.Append(" <tr height=24 style='mso-height-source:userset;height:22.5pt'>                                                                                    \n");
                htmlStr.Append("  <td height=24 class=xl8415939 style='height:22.5pt;border-bottom:1pt solid windowtext;border-top:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                  \n");
                htmlStr.Append("  <td style='height:22.5pt;border-bottom:1pt solid windowtext;border-top:none;border-left:none;border-right:none' class=xl7315939>&nbsp;</td>                                                                                                                  \n");
                htmlStr.Append("  <td class=xl8215939 width=127 style='height:22.5pt;border-bottom:1pt solid windowtext;border-top:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                     \n");
                htmlStr.Append("  <td colspan=7 class=xl156159391 width=347 style='height:22.5pt;border-bottom:1pt solid windowtext;border-top:none;border-left:none;border-right:none'>Ngày <font                                                                          \n");
                htmlStr.Append("  class='font1015939'>(Date)</font><font class='font915939'> </font><font                                                                          \n");
                htmlStr.Append("  class='font615939'><span style='mso-spacerun:yes'>&nbsp;    " + dt.Rows[0]["invoiceissueddate_dd"] + "  </span>&nbsp;&nbsp;tháng </font><font                                                                 \n");
                htmlStr.Append("  class='font1015939'>(month)</font><font class='font615939'><span                                                                                 \n");
                htmlStr.Append("  style='mso-spacerun:yes'>&nbsp;&nbsp; " + dt.Rows[0]["invoiceissueddate_mm"] + " </span>&nbsp;&nbsp;năm </font><font                                                                                   \n");
                htmlStr.Append("  class='font1015939'>(year)</font><font class='font615939'><span                                                                                  \n");
                htmlStr.Append("  style='mso-spacerun:yes'>&nbsp;&nbsp;" + dt.Rows[0]["invoiceissueddate_yyyy"] + "</br>" + v_titlePageNumber + " </span></font> </td>                                                                                                    \n");
                htmlStr.Append("  <td class=xl8015939 style='height:22.5pt;border-bottom:1pt solid windowtext;border-top:none;border-left:none;border-right:none' colspan=4>S&#7889; (<font class='font1015939'>No</font><font                                                                 \n");
                htmlStr.Append("  class='font815939'>.):</font><font class='font715939'><span                                                                                      \n");
                htmlStr.Append("  style='mso-spacerun:yes'>      </span></font><font class='font1115939'><span                                                                     \n");
                htmlStr.Append("  style='mso-spacerun:yes'> </span> " + dt.Rows[0]["InvoiceNumber"] + "</font></td>                                                                                             \n");
                htmlStr.Append("  <td class=xl8315939 style='height:22.5pt;border-bottom:1pt solid windowtext;border-top:none;border-left:none;border-right:none' width=12 style='width:9pt'>&nbsp;</td>                                                                                       \n");
                htmlStr.Append(" </tr>                                                                                                                                             \n");
                htmlStr.Append(" <tr height=24 style='mso-height-source:userset;height:22.5pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>                                                                                    \n");
                htmlStr.Append("  <td height=24 class=xl8415939 style='height:22.5pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                  \n");
                htmlStr.Append("  <td style='border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl8715939 colspan=10>Đơn vị bán hàng (<font                                                                            \n");
                htmlStr.Append("  class='font1015939'>Company name</font><font class='font615939'>): </font><font                                                                  \n");
                htmlStr.Append("  class='font515939'>" + dt.Rows[0]["Seller_Name"] + "</font></td>                                                                               \n");
                htmlStr.Append("  <td style='border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl7315939>&nbsp;</td>                                                                                                                  \n");
                htmlStr.Append("  <td style='border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl7315939>&nbsp;</td>                                                                                                                  \n");
                htmlStr.Append("  <td style='border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl7315939>&nbsp;</td>                                                                                                                  \n");
                htmlStr.Append("  <td style='border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl8115939>&nbsp;</td>                                                                                                                  \n");
                htmlStr.Append(" </tr>                                                                                                                                             \n");
                htmlStr.Append(" <tr height=24 style='mso-height-source:userset;height:22.5pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>                                                                                    \n");
                htmlStr.Append("  <td height=24 class=xl8415939 style='height:22.5pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                  \n");
                htmlStr.Append("  <td class=xl8715939 style='border-top:none;border-bottom:none;border-left:none;border-right:none' colspan=13>Điạ chỉ (<font                                                                                   \n");
                htmlStr.Append("  class='font1015939'>Address</font><font class='font615939'>): " + dt.Rows[0]["Seller_Address"] + "</font></td>                                                                                                                  \n");
                htmlStr.Append("  <td style='border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl8115939>&nbsp;</td>                                                                                                                  \n");
                htmlStr.Append(" </tr>                                                                                                                                             \n");
                htmlStr.Append(" <tr height=24 style='mso-height-source:userset;height:22.5pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>                                                                                    \n");
                htmlStr.Append("  <td height=24 class=xl8415939 style='height:22.5pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                  \n");
                htmlStr.Append("  <td style='border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl8715939 colspan=5>Mã số thuế (<font                                                                                      \n");
                htmlStr.Append("  class='font1015939'>Tax code</font><font class='font615939'>): </font><font                                                                      \n");
                htmlStr.Append("  class='font515939'>" + dt.Rows[0]["Seller_TaxCode"] + "</font></td>                                                                                               \n");
                htmlStr.Append("  <td style='border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl8715939>&nbsp;</td>                                                                                                                  \n");
                htmlStr.Append("  <td style='border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl8715939>&nbsp;</td>                                                                                                                  \n");
                htmlStr.Append("  <td style='border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl7315939>&nbsp;</td>                                                                                                                  \n");
                htmlStr.Append("  <td style='border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl7315939>&nbsp;</td>                                                                                                                  \n");
                htmlStr.Append("  <td style='border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl7315939>&nbsp;</td>                                                                                                                  \n");
                htmlStr.Append("  <td style='border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl7315939>&nbsp;</td>                                                                                                                  \n");
                htmlStr.Append("  <td style='border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl7315939>&nbsp;</td>                                                                                                                  \n");
                htmlStr.Append("  <td style='border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl7315939>&nbsp;</td>                                                                                                                  \n");
                htmlStr.Append("  <td style='border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl8115939>&nbsp;</td>                                                                                                                  \n");
                htmlStr.Append(" </tr>                                                                                                                                             \n");
                htmlStr.Append(" <tr height=24 style='mso-height-source:userset;height:22.5pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>                                                                                    \n");
                htmlStr.Append("  <td height=24 class=xl8415939 style='height:22.5pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                  \n");
                htmlStr.Append("  <td style='border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl8715939 colspan=7>Điện thoại (<font                                                                                                   \n");
                htmlStr.Append("  class='font1015939'>Tel</font><font class='font615939'>): " + dt.Rows[0]["Seller_Tel"] + "-                                                                     \n");
                htmlStr.Append("  Fax: " + dt.Rows[0]["Seller_Fax"] + "</font></td>                                                                                                                 \n");
                htmlStr.Append("  <td style='border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl7315939>&nbsp;</td>                                                                                                                  \n");
                htmlStr.Append("  <td style='border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl7315939>&nbsp;</td>                                                                                                                  \n");
                htmlStr.Append("  <td style='border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl7315939>&nbsp;</td>                                                                                                                  \n");
                htmlStr.Append("  <td style='border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl7315939>&nbsp;</td>                                                                                                                  \n");
                htmlStr.Append("  <td style='border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl7315939>&nbsp;</td>                                                                                                                  \n");
                htmlStr.Append("  <td style='border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl7315939>&nbsp;</td>                                                                                                                  \n");
                htmlStr.Append("  <td style='border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl8115939>&nbsp;</td>                                                                                                                  \n");
                htmlStr.Append(" </tr>                                                                                                                                             \n");
                htmlStr.Append(" <tr height=24 style='mso-height-source:userset;height:22.5pt;border-bottom:1pt solid windowtext;border-top:none;border-left:none;border-right:none'>                                                                                    \n");
                htmlStr.Append("  <td height=24 class=xl8415939 style='height:22.5pt;border-bottom:1pt solid windowtext;border-top:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                  \n");
                htmlStr.Append("  <td class=xl8715939 style='height:22.5pt;border-bottom:1pt solid windowtext;border-top:none;border-left:none;border-right:none' colspan=13>Số tài khoản (<font                                                                                                 \n");
                htmlStr.Append("  class='font1015939'>Account number</font><font class='font615939'>) : " + dt.Rows[0]["seller_accountno"] + "</font></td>                    \n");
                htmlStr.Append("  <td class=xl8115939 style='height:22.5pt;border-bottom:1pt solid windowtext;border-top:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                                                  \n");
                htmlStr.Append(" </tr>                                                                                                                                             \n");
                htmlStr.Append(" <tr height=24 style='mso-height-source:userset;height:22.5pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>                                                                                    \n");
                htmlStr.Append("  <td height=24 class=xl8415939 style='height:22.5pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                  \n");
                htmlStr.Append("  <td class=xl8715939 style='border-right:none;border-top:none;border-bottom:none;border-left:none;border-right:none' colspan=4>Họ tên người mua hàng (<font                                                                                        \n");
                htmlStr.Append("  class='font1015939'>Customer's name</font><span style='display:'><font                                                                       \n");
                htmlStr.Append("  class='font615939'>):</font></span></td>                                                                                                         \n");
                htmlStr.Append("  <td colspan=10 class=xl9215939 style='border-right:none;border-top:none;border-bottom:none;border-left:none;border-right:none'>" + dt.Rows[0]["Buyer"] + "</td>                                                                 \n");
                htmlStr.Append(" </tr>                                                                                                                                             \n");
                htmlStr.Append(" <tr height=24 style='mso-height-source:userset;height:22.5pt;border-right:none;border-top:none;border-bottom:none;border-left:none;border-right:none'>                                                                                    \n");
                htmlStr.Append("  <td height=24 class=xl8415939 style='height:22.5pt;border-right:none;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                  \n");
                htmlStr.Append("  <td class=xl8715939 style='border-right:none;border-top:none;border-bottom:none;border-left:none;border-right:none' colspan=3>Tên đơn vị (<font                                                                                                   \n");
                htmlStr.Append("  class='font1015939'>Company's name):</font></td>                                                                                                 \n");
                htmlStr.Append("  <td colspan=10 style='border-right:none;border-top:none;border-bottom:none;border-left:none;border-right:none;' class=xl9115939>" + dt.Rows[0]["BuyerLegalName"] + "</td>                                                                                                                  \n");
                htmlStr.Append("  <td style='border-right:none;border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl9315939>&nbsp;</td>                                                                                                                  \n");
                htmlStr.Append(" </tr>                                                                                                                                             \n");
                htmlStr.Append(" <tr height=24 style='mso-height-source:userset;height:22.5pt'>                                                                                    \n");
                htmlStr.Append("  <td height=24 class=xl8415939 style='height:22.5pt'>&nbsp;</td>                                                                                  \n");
                htmlStr.Append("  <td class=xl8715939 style='border-right:none;border-top:none;border-bottom:none;border-left:none;border-right:none' colspan=2>Địa chỉ (<font                                                                                                      \n");
                htmlStr.Append("  class='font1015939'>Address</font><font class='font615939'>):</td>                                                                                                 \n");
                htmlStr.Append("  <td class=xl8715939 style='border-right:none;border-top:none;border-bottom:none;border-left:none;border-right:none;font-size:12.5pt' colspan=11>" + dt.Rows[0]["BuyerAddress"] + "    </td>                                                                                                 \n");
                htmlStr.Append("  <td style='border-right:none;border-top:none;border-bottom:none;border-left:none;border-right:none' class=xl9415939>&nbsp;</td>                                                                                                                  \n");
                htmlStr.Append(" </tr>                                                                                                                                             \n");
                htmlStr.Append(" <tr height=24 style='mso-height-source:userset;height:22.5pt;border-top:none;border-left:none;border-right:none'>                                                                                    \n");
                htmlStr.Append("  <td height=24  class=xl8415939 style='height:22.5pt;border-top:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                  \n");
                htmlStr.Append("  <td class=xl8715939 colspan=5 style='height:22.5pt;border-top:none;border-left:none;border-right:none'>Mã số thuế (<font                                                                                         \n");
                htmlStr.Append("  class='font1015939'>Tax code</font><font class='font515939'>):<span                                                                          \n");
                htmlStr.Append("  style='mso-spacerun:yes'> </span></font>" + dt.Rows[0]["BuyerTaxCode"] + "</td>                                                                                                    \n");
                htmlStr.Append("  <td class=xl8715939 colspan=8 style='height:22.5pt;border-top:none;border-left:none;border-right:none'>Hình thức thanh toán (<font                                                                                                 \n");
                htmlStr.Append("  class='font1015939'>Payment term</font><font class='font615939'>):  </font>" + dt.Rows[0]["PaymentMethodCK"] + "</td>                                                                 \n");
                htmlStr.Append("  <td class=xl8115939 style='height:22.5pt;border-top:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                                                  \n");
                htmlStr.Append(" </tr>                                                                                                                                             \n");
                htmlStr.Append(" <tr height=24 style='mso-height-source:userset;height:22.5pt;border-top:none;border-left:none;border-right:none'>                                                                                    \n");
                htmlStr.Append("  <td height=24  class=xl8415939 style='height:22.5pt;border-top:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                  \n");
                htmlStr.Append("  <td class=xl8715939 colspan=5 style='height:22.5pt;border-top:none;border-left:none;border-right:none'>Loại tiền tệ (<font                                                                                         \n");
                htmlStr.Append("  class='font1015939'>Currency</font>):  <span                                                                          \n");
                htmlStr.Append("  style='mso-spacerun:yes'>" + dt.Rows[0]["CurrencyCodeUSD"] + " </span></td>                                                                                                    \n");
                htmlStr.Append("  <td class=xl8715939 colspan=8 style='height:22.5pt;border-top:none;border-left:none;border-right:none'>Tỷ giá (<font                                                                                                 \n");
                htmlStr.Append("  class='font1015939'>Exchagne rate</font><font class='font615939'>):  </font></td>                                                                 \n");
                htmlStr.Append("  <td class=xl8115939 style='height:22.5pt;border-top:none;border-left:none;border-right:none'>&nbsp;" + dt.Rows[0]["ExchangeRate"] + "</td>                                                                                                                  \n");
                htmlStr.Append(" </tr>                                                                                                                                             \n");
                htmlStr.Append(" <tr height=24 style='mso-height-source:userset;height:22.5pt;border-top:none;border-left:none;border-right:none'>                                                                                    \n");
                htmlStr.Append("  <td height=24  class=xl8415939 style='height:22.5pt;border-top:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                  \n");
                htmlStr.Append("  <td class=xl8715939 colspan=5 style='height:22.5pt;border-top:none;border-left:none;border-right:none'>Số tờ khai (<font                                                                                         \n");
                htmlStr.Append("  class='font1015939'>Declaration no</font><font class='font515939'>):<span                                                                          \n");
                htmlStr.Append("  style='mso-spacerun:yes'>" + dt.Rows[0]["DECLARE_NO"] + " </span></font></td>                                                                                                    \n");
                htmlStr.Append("  <td class=xl8715939 colspan=8 style='height:22.5pt;border-top:none;border-left:none;border-right:none'>Ngày tờ khai (<font                                                                                                 \n");
                htmlStr.Append("  class='font1015939'>Declaration date</font><font class='font615939'>):  </font>" + dt.Rows[0]["DECLARE_DT"] + "</td>                                                                 \n");
                htmlStr.Append("  <td class=xl8115939 style='height:22.5pt;border-top:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                                                  \n");
                htmlStr.Append(" </tr>                                                                                                                                             \n");
                
                htmlStr.Append(" <tr class=xl10215939 height=21 style='height:19.5pt;border-right:1pt solid windowtext;'>                                                                                             \n");
                htmlStr.Append("  <td colspan=2 height=21 class=xl15915939 style='height:19.5pt;;border-right:1pt solid windowtext;'>STT</td>                                                                          \n");
                htmlStr.Append("  <td colspan=4 class=xl15915939 style='border-right:.5pt solid black;border-right:1pt solid windowtext;'>Tên hàng                                                                    \n");
                htmlStr.Append("  hóa, dịch vụ</td>                                                                                                                    \n");
                htmlStr.Append("  <td class=xl10115939 style='border-bottom:none;border-right:1pt solid windowtext;'>Đơn vị tính</td>                                                                    \n");
                htmlStr.Append("  <td colspan=2 class=xl15915939 style='border-right:1pt solid windowtext;'>Số                                                                    \n");
                htmlStr.Append("  lượng</td>                                                                                                                            \n");
                htmlStr.Append("  <td colspan=3 style='border-right:1pt solid black;border-bottom:none;' class=xl10115939>Đơn giá</td>                                                                                            \n");
                htmlStr.Append("  <td colspan=3 class=xl15915939 style='border-right:1pt solid black;'>Thành                                                                       \n");
                htmlStr.Append("  ti&#7873;n</td>                                                                                                                                  \n");
                htmlStr.Append(" </tr>                                                                                                                                             \n");
                htmlStr.Append(" <tr class=xl10315939 height=18 style='height:16.5‬pt;border-right:1pt solid black'>                                                                                             \n");
                htmlStr.Append("  <td colspan=2 height=18 class=xl16115939 style='height:16.5‬pt;border-right:1pt solid black;border-bottom:1pt solid windowtext'>No.</td>                                                                          \n");
                htmlStr.Append("  <td colspan=4 class=xl16115939 style='border-bottom:1pt solid windowtext;border-right:1pt solid black'>Description</td>                                                            \n");
                htmlStr.Append("  <td style='border-bottom:1pt solid windowtext;border-right:1pt solid black;border-top:none;' class=xl10315939>Unit</td>                                                                                                                   \n");
                htmlStr.Append("  <td colspan=2 class=xl16115939 style='border-bottom:1pt solid windowtext;border-right:1pt solid black'>Quantity</td>                                                                \n");
                htmlStr.Append("  <td style='border-bottom:none;border-right:1pt solid black;border-bottom:1pt solid windowtext;border-top:none;' colspan=3  class=xl10315939>Unit price</td>                                                                                                   \n");
                htmlStr.Append("  <td colspan=3 class=xl16115939 style='border-right:.5pt solid black;border-bottom:1pt solid windowtext'>Amount</td>                                                                 \n");
                htmlStr.Append(" </tr>                                                                                                                                             \n");
                htmlStr.Append(" <tr class=xl10515939 height=20 style='mso-height-source:userset;height:18.75pt‬;border-right:1pt solid black;border-bottom:1pt solid black'>                                                                   \n");
                htmlStr.Append("  <td colspan=2 height=20 class=xl16315939 style='height:18.75pt‬;border-right:1pt solid black;border-bottom:1pt solid black'>1</td>                                                                            \n");
                htmlStr.Append("  <td colspan=4 class=xl16315939 style='border-right:1pt solid black;border-bottom:1pt solid black'>2</td>                                                                      \n");
                htmlStr.Append("  <td class=xl10415939 style='border-right:1pt solid black;border-bottom:1pt solid black'>3</td>                                                                                                                      \n");
                htmlStr.Append("  <td colspan=2 class=xl16315939 style='border-right:1pt solid black'>4</td>                                                                      \n");
                htmlStr.Append("  <td colspan=3 class=xl10415939 style='border-right:1pt solid black;border-bottom:1pt solid black'>5</td>                                                                                                            \n");
                htmlStr.Append("  <td colspan=3 class=xl16315939 style='border-right:1pt solid black;border-bottom:1pt solid black'>6</td>                                                                      \n");
                htmlStr.Append(" </tr>                                                                                                                                             \n");


                v_rowHeight = "30.0pt"; //"26.5pt";
                v_rowHeightEmpty = "22.0pt";
                v_rowHeightNumber = 26;

                v_rowHeightLast = "26.0pt";// "23.5pt";
                v_rowHeightLastNumber = 26;// 23.5;
                v_rowHeightEmptyLast = "23.5pt"; //"23.5pt";


                for (int dtR = 0; dtR < page[k]; dtR++)
                {
                    if (!vlongItemName && dt_d.Rows[v_index]["ITEM_NAME"].ToString().Length >= 92)
                    {
                        v_rowHeight = "30.0pt"; //"26.5pt";    
                        v_rowHeightLast = "26.0pt"; //"27.5pt";
                        v_rowHeightLastNumber = 26;//27.5;
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

                        htmlStr.Append(" <tr class=xl11315939 height=22 style='mso-height-source:userset;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "‬'>                                                                   \n");

                        if (Double.Parse(dt.Rows[0]["INVOICE_DATE_110"].ToString()) <= 20230512)
                        {
                            htmlStr.Append(" 	<td colspan=2 class=xl16615939  height=24 width=39 style='border-right:1pt solid black;border-bottom:none             \n");
                            htmlStr.Append(" 	  height:22.5pt;width:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "' align=left valign=top><![if !vml]><span style='mso-ignore:vglayout;                                                                   \n");
                            htmlStr.Append(" 	  position:absolute;z-index:3;margin-left:-18px;margin-top:85px;width:875px;           \n");
                            htmlStr.Append(" 	  height:202px'><img width=875 height=202                                                                 \n");
                            htmlStr.Append(" 	  src='D:/webproject/e-invoice-ws/02.Web/EInvoice/img/Suheung_4.png'             \n");
                            htmlStr.Append(" 	  v:shapes='Picture_x0020_14'></span><![endif]><span style='mso-ignore:vglayout2'>                                                                    \n");
                            htmlStr.Append(" 	  <table cellpadding=0 cellspacing=0>             \n");
                            htmlStr.Append(" 	   <tr>                                                                     \n");
                            htmlStr.Append(" 		<td height=22 class=xl11215939 style='height:21.375pt‬;border-bottom:none;border-top:none'>&nbsp;&nbsp;&nbsp;&nbsp;" + dt_d.Rows[v_index][7] + "</td>             \n");
                            htmlStr.Append(" 	   </tr>             \n");
                            htmlStr.Append(" 	  </table>             \n");
                            htmlStr.Append(" 	  </span></td>             \n");

                        }
                        else
                        {
                            htmlStr.Append(" 	<td colspan=2 class=xl16615939  height=24 width=39 style='border-right:1pt solid black;border-bottom:none             \n");
                            htmlStr.Append(" 	  height:22.5pt;width:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "' align=left valign=top><![if !vml]><span style='mso-ignore:vglayout;                                                                   \n");
                            htmlStr.Append(" 	  position:absolute;z-index:3;margin-left:-2px;margin-top:85px;width:875px;           \n");
                            htmlStr.Append(" 	  height:202px'><img width=875 height=202                                                                 \n");
                            htmlStr.Append(" 	  src='D:/webproject/e-invoice-ws/02.Web/EInvoice/img/Suheung_2_20230512.png'             \n");
                            htmlStr.Append(" 	  v:shapes='Picture_x0020_14'></span><![endif]><span style='mso-ignore:vglayout2'>                                                                    \n");
                            htmlStr.Append(" 	  <table cellpadding=0 cellspacing=0>             \n");
                            htmlStr.Append(" 	   <tr>                                                                     \n");
                            htmlStr.Append(" 		<td height=22 class=xl11215939 style='height:21.375pt‬;border-bottom:none;border-top:none'>&nbsp;&nbsp;&nbsp;&nbsp;" + dt_d.Rows[v_index][7] + "</td>             \n");
                            htmlStr.Append(" 	   </tr>             \n");
                            htmlStr.Append(" 	  </table>             \n");
                            htmlStr.Append(" 	  </span></td>             \n");


                        }


                        htmlStr.Append("  <td colspan=4 rowspan=1 class=xl16915939 width=264 style='border-right:1pt solid black;border-bottom:none                                                         \n");
                        htmlStr.Append("  width:198pt'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                                                                                      \n");
                        htmlStr.Append("  <td class=xl10615939 style='border-left:none;border-bottom:none;border-right:1pt solid black '>" + dt_d.Rows[v_index][1] + "&nbsp;</td>                                                                                        \n");
                        htmlStr.Append("  <td class=xl10715939 >&nbsp;" + dt_d.Rows[v_index][2] + "</td>                                                                                        \n");
                        htmlStr.Append("  <td class=xl10815939 style='border-left:none;border-right:1pt solid black'>&nbsp;</td>                                                                                                                 \n");
                        htmlStr.Append("  <td class=xl10915939 style='border-left:none;'colspan=2>&nbsp;" + dt_d.Rows[v_index][3] + "</td>                                                                                        \n");
                        // htmlStr.Append("  <td class=xl11015939>&nbsp;</td>                                                                                                                 \n");
                        htmlStr.Append("  <td class=xl11115939 style='border-left:none;border-right:1pt solid black'>&nbsp;</td>                                                                                                                 \n");
                        htmlStr.Append("  <td class=xl11215939 style='border-left:none'>&nbsp;</td>                                                                                        \n");
                        htmlStr.Append("  <td class=xl11315939 style='border-right:none;border-bottom:none;border-left:none'>" + dt_d.Rows[v_index][4] + "&nbsp;</td>                                                                                                                 \n");
                        htmlStr.Append("  <td class=xl11115939><span style='mso-spacerun:yes'>   </span></td>                                                                              \n");
                        htmlStr.Append(" </tr> \n");


                    }
                    else if (dtR == page[k] - 1)//dong cuoi moi trang
                    {
                        if (k < v_countNumberOfPages - 1) //trang giua
                        {
                            htmlStr.Append(" <tr class=xl8715939 height=24 style='mso-height-source:userset;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                    \n");
                            htmlStr.Append("  <td colspan=2 height=24 class=xl16615939 width=39 style='border-right:1pt solid black;border-bottom:1pt solid black;                                                          \n");
                            htmlStr.Append("  height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ";width:29pt'>&nbsp;" + dt_d.Rows[v_index][7] + "</td>                                                                                                            \n");
                            htmlStr.Append("  <td colspan=4 class=xl15615939 width=264 style='width:198pt;border-right:1pt solid black;border-bottom:1pt solid black;border-top:none; '>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                                         \n");
                            htmlStr.Append("  <td class=xl12115939 width=77 style='width:72.5‬pt;border-right:1pt solid black;border-bottom:1pt solid black; '>&nbsp;" + dt_d.Rows[v_index][1] + "</td>                                                                                     \n");
                            htmlStr.Append("  <td class=xl12215939 width=78 style='border-left:none;width:72.5‬pt;border-right:none;border-bottom:1pt solid black; '>&nbsp;" + dt_d.Rows[v_index][2] + "</td>                                                                    \n");
                            htmlStr.Append("  <td class=xl12315939 width=6 style='width:5pt;border-right:1pt solid black;border-bottom:1pt solid black; '>&nbsp;</td>                                                                                       \n");
                            htmlStr.Append("  <td colspan=2 class=xl11615939 style='border-left:none;border-right:none;border-bottom:1pt solid black; '>&nbsp;" + dt_d.Rows[v_index][3] + "</td>                                                                              \n");
                            htmlStr.Append("  <td class=xl12415939 width=6 style='width:5pt;border-right:1pt solid black;border-bottom:1pt solid black; '>&nbsp;</td>                                                                                       \n");
                            htmlStr.Append("  <td colspan=2 class=xl11615939 style='border-left:none;border-right:none;border-bottom:1pt solid black; '>&nbsp;" + dt_d.Rows[v_index][4] + "</td>                                                                              \n");
                            htmlStr.Append("  <td class=xl12515939 width=12 style='width:9pt;border-right:1pt solid black;border-bottom:1pt solid black; '>&nbsp;</td>                                                                                      \n");
                            htmlStr.Append(" </tr>                                                                                                                                             \n");

                        }
                        else // trang cuoi
                        {
                            if (dtR == rowsPerPage - 1) // du 11 dong
                            {
                                htmlStr.Append(" <tr class=xl8715939 height=24 style='mso-height-source:userset;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                    \n");
                                htmlStr.Append("  <td colspan=2 height=24 class=xl16615939 width=39 style='border-right:1pt solid black;border-bottom:1pt solid black;                                                          \n");
                                htmlStr.Append("  height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ";width:29pt'>&nbsp;" + dt_d.Rows[v_index][7] + "</td>                                                                                                            \n");
                                htmlStr.Append("  <td colspan=4 class=xl15615939 width=264 style='width:198pt;border-right:1pt solid black;border-bottom:1pt solid black;border-top:none; '>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                                         \n");
                                htmlStr.Append("  <td class=xl12115939 width=77 style='width:72.5‬pt;border-right:1pt solid black;border-bottom:1pt solid black; '>&nbsp;" + dt_d.Rows[v_index][1] + "</td>                                                                                     \n");
                                htmlStr.Append("  <td class=xl12215939 width=78 style='border-left:none;width:72.5‬pt;border-right:none;border-bottom:1pt solid black; '>&nbsp;" + dt_d.Rows[v_index][2] + "</td>                                                                    \n");
                                htmlStr.Append("  <td class=xl12315939 width=6 style='width:5pt;border-right:1pt solid black;border-bottom:1pt solid black; '>&nbsp;</td>                                                                                       \n");
                                htmlStr.Append("  <td colspan=2 class=xl11615939 style='border-left:none;border-right:none;border-bottom:1pt solid black; '>&nbsp;" + dt_d.Rows[v_index][3] + "</td>                                                                              \n");
                                htmlStr.Append("  <td class=xl12415939 width=6 style='width:5pt;border-right:1pt solid black;border-bottom:1pt solid black; '>&nbsp;</td>                                                                                       \n");
                                htmlStr.Append("  <td colspan=2 class=xl11615939 style='border-left:none;border-right:none;border-bottom:1pt solid black; '>&nbsp;" + dt_d.Rows[v_index][4] + "</td>                                                                              \n");
                                htmlStr.Append("  <td class=xl12515939 width=12 style='width:9pt;border-right:1pt solid black;border-bottom:1pt solid black; '>&nbsp;</td>                                                                                      \n");
                                htmlStr.Append(" </tr>                                                                                                                                             \n");
                            }
                            else
                            {
                                htmlStr.Append(" <tr class=xl8715939 height=24 style='mso-height-source:userset;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                    \n");
                                htmlStr.Append("  <td colspan=2 height=24 class=xl16615939 width=39 style='border-right:1pt solid black;border-bottom:1pt solid black;                                                          \n");
                                htmlStr.Append("  height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ";width:29pt'>&nbsp;" + dt_d.Rows[v_index][7] + "</td>                                                                                                            \n");
                                htmlStr.Append("  <td colspan=4 class=xl15615939 width=264 style='width:198pt;border-right:1pt solid black;border-bottom:1pt solid black;border-top:none; '>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                                         \n");
                                htmlStr.Append("  <td class=xl12115939 width=77 style='width:72.5‬pt;border-right:1pt solid black;border-bottom:1pt solid black; '>&nbsp;" + dt_d.Rows[v_index][1] + "</td>                                                                                     \n");
                                htmlStr.Append("  <td class=xl12215939 width=78 style='border-left:none;width:72.5‬pt;border-right:none;border-bottom:1pt solid black; '>&nbsp;" + dt_d.Rows[v_index][2] + "</td>                                                                    \n");
                                htmlStr.Append("  <td class=xl12315939 width=6 style='width:5pt;border-right:1pt solid black;border-bottom:1pt solid black; '>&nbsp;</td>                                                                                       \n");
                                htmlStr.Append("  <td colspan=2 class=xl11615939 style='border-left:none;border-right:none;border-bottom:1pt solid black; '>&nbsp;" + dt_d.Rows[v_index][3] + "</td>                                                                              \n");
                                htmlStr.Append("  <td class=xl12415939 width=6 style='width:5pt;border-right:1pt solid black;border-bottom:1pt solid black; '>&nbsp;</td>                                                                                       \n");
                                htmlStr.Append("  <td colspan=2 class=xl11615939 style='border-left:none;border-right:none;border-bottom:1pt solid black; '>&nbsp;" + dt_d.Rows[v_index][4] + "</td>                                                                              \n");
                                htmlStr.Append("  <td class=xl12515939 width=12 style='width:9pt;border-right:1pt solid black;border-bottom:1pt solid black; '>&nbsp;</td>                                                                                      \n");
                                htmlStr.Append(" </tr>                                                                                                                                             \n");
                            }

                        }
                    }
                    else
                    { // dong giua                                                                                                                                    
                        htmlStr.Append(" <tr class=xl10215939 height=24 style='mso-height-source:userset;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                   \n");
                        htmlStr.Append("  <td colspan=2 height=24 class=xl16615939 width=39 style='border-right:1pt solid black;border-top:none; border-bottom:none                                                         \n");
                        htmlStr.Append("  height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ";width:29pt'>&nbsp;" + dt_d.Rows[v_index][7] + "</td>                                                                                                            \n");
                        htmlStr.Append("  <td colspan=4 class=xl15615939 width=264 style='width:198pt;border-right:1pt solid black;border-bottom:none;border-top:none;'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                                         \n");
                        htmlStr.Append("  <td class=xl11815939  style='width:198pt;border-right:1pt solid black;border-bottom:none'>&nbsp;" + dt_d.Rows[v_index][1] + "</td>                                                                                                                 \n");
                        htmlStr.Append("  <td class=xl11615939 style='border-left:none;border-bottom:none'>&nbsp;" + dt_d.Rows[v_index][2] + "</td>                                                                                        \n");
                        htmlStr.Append("  <td class=xl11515939 style='border-right:1pt solid black;border-bottom:none'>&nbsp;</td>                                                                                                                 \n");
                        htmlStr.Append("  <td colspan=2 class=xl11615939 style='border-left:none;border-bottom:none'>&nbsp;" + dt_d.Rows[v_index][3] + "</td>                                                               \n");
                        htmlStr.Append("  <td class=xl11715939 style='border-right:1pt solid black;border-bottom:none'>&nbsp;</td>                                                                                                                 \n");
                        htmlStr.Append("  <td colspan=2 class=xl11615939 style='border-left:none;border-bottom:none'>&nbsp;" + dt_d.Rows[v_index][4] + "</td>                                                                              \n");
                        htmlStr.Append("  <td class=xl11715939 style='border-right:1pt solid black;border-bottom:none'>&nbsp;</td>                                                                                                                 \n");
                        htmlStr.Append(" </tr>                                                                                                                                       \n");

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
                            htmlStr.Append(" <tr class=xl8715939 height=24 style='mso-height-source:userset;height:" + v_rowHeightEmptyLast + "'>                                                                    \n");
                            htmlStr.Append("  <td colspan=2 height=24 class=xl16615939 width=39 style='border-right:.5pt solid black;  border-bottom:1pt solid black;                                                         \n");
                            htmlStr.Append("  height:" + v_rowHeightEmptyLast + ";width:29pt'>&nbsp;</td>                                                                                                            \n");
                            htmlStr.Append("  <td colspan=4 class=xl15615939 width=264 style='width:5pt;border-right:1pt solid black;border-bottom:1pt solid black;border-top:none  '>&nbsp;</td>                                                                         \n");
                            htmlStr.Append("  <td class=xl12115939 width=77 style='width:5pt;border-right:1pt solid black;border-bottom:1pt solid black; '>&nbsp;</td>                                                                                     \n");
                            htmlStr.Append("  <td class=xl12215939 width=78 style='width:5pt;border-right:none;border-bottom:1pt solid black; '>&nbsp;</td>                                                                    \n");
                            htmlStr.Append("  <td class=xl12315939 width=6 style='width:5pt;border-right:1pt solid black;border-bottom:1pt solid black; '>&nbsp;</td>                                                                                       \n");
                            htmlStr.Append("  <td colspan=2 class=xl11615939 style='width:5pt;border-right:none;border-bottom:1pt solid black; '>&nbsp;</td>                                                                              \n");
                            htmlStr.Append("  <td class=xl12415939 width=6 style='width:5pt;border-right:1pt solid black;border-bottom:1pt solid black; '>&nbsp;</td>                                                                                       \n");
                            htmlStr.Append("  <td colspan=2 class=xl11615939 style='border-left:none;border-right:none;border-bottom:1pt solid black; '>&nbsp;</td>                                                                              \n");
                            htmlStr.Append("  <td class=xl12515939 width=12 style='width:9pt;border-right:1pt solid black;border-bottom:1pt solid black; '>&nbsp;</td>                                                                                      \n");
                            htmlStr.Append(" </tr>                                                                                                                                             \n");


                        }
                        else
                        {
                            htmlStr.Append(" <tr class=xl10215939 height=24 style='mso-height-source:userset;height:" + v_rowHeightEmptyLast + "'>                                                                   \n");
                            htmlStr.Append("  <td colspan=2 height=24 class=xl16615939 width=39 style='border-top:none;border-bottom:none;border-left:none;border-right:1pt solid black                                                       \n");
                            htmlStr.Append("  height:" + v_rowHeightEmptyLast + ";width:29pt'>&nbsp;</td>                                                                                                            \n");
                            htmlStr.Append("  <td colspan=4 class=xl15615939 width=264 style='border-top:none;border-bottom:none;border-left:1pt solid black;border-right:1pt solid black'>&nbsp;</td>                                                                         \n");
                            htmlStr.Append("  <td class=xl11815939 style='border-top:none;border-bottom:none;border-left:none;border-right:1pt solid black'>&nbsp;</td>                                                                                                                 \n");
                            htmlStr.Append("  <td class=xl11615939 style='border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                        \n");
                            htmlStr.Append("  <td class=xl11515939 style='border-top:none;border-bottom:none;border-left:none;border-right:1pt solid black'>&nbsp;</td>                                                                                                                 \n");
                            htmlStr.Append("  <td colspan=2 class=xl11615939 style='border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                              \n");
                            htmlStr.Append("  <td class=xl11715939 style='border-top:none;border-bottom:none;border-left:none;border-right:1pt solid black'>&nbsp;</td>                                                                                                                 \n");
                            htmlStr.Append("  <td colspan=2 class=xl11615939 style='border-top:none;border-bottom:none;border-left:none;border-rightnone'>&nbsp;</td>                                                                              \n");
                            htmlStr.Append("  <td class=xl11715939 style='border-top:none;border-bottom:none;border-left:none;border-right:1pt solid black'>&nbsp;</td>                                                                                                                 \n");
                            htmlStr.Append(" </tr>                                                                                                                                             \n");

                        }
                    } // for

                }//Trang cuoi 11 dong

                if (k < v_countNumberOfPages - 1)
                {
                    htmlStr.Append(" <tr class=xl10215939 height=24 style='mso-height-source:userset;height:" + (v_spacePerPage).ToString() + "pt'>                                                                   \n");
                    htmlStr.Append("  <td colspan=2 height=24 class=xl16615939 width=39 style='border-top:none;border-bottom:none;border-left:none;border-right:1pt solid black                                                       \n");
                    htmlStr.Append("  height:" + (v_spacePerPage).ToString() + "pt;width:29pt'>&nbsp;</td>                                                                                                            \n");
                    htmlStr.Append("  <td colspan=4 class=xl15615939 width=264 style='border-top:none;border-bottom:none;border-left:1pt solid black;border-right:1pt solid black'>&nbsp;</td>                                                                         \n");
                    htmlStr.Append("  <td class=xl11815939 style='border-top:none;border-bottom:none;border-left:none;border-right:1pt solid black'>&nbsp;</td>                                                                                                                 \n");
                    htmlStr.Append("  <td class=xl11615939 style='border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                        \n");
                    htmlStr.Append("  <td class=xl11515939 style='border-top:none;border-bottom:none;border-left:none;border-right:1pt solid black'>&nbsp;</td>                                                                                                                 \n");
                    htmlStr.Append("  <td colspan=2 class=xl11615939 style='border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                              \n");
                    htmlStr.Append("  <td class=xl11715939 style='border-top:none;border-bottom:none;border-left:none;border-right:1pt solid black'>&nbsp;</td>                                                                                                                 \n");
                    htmlStr.Append("  <td colspan=2 class=xl11615939 style='border-top:none;border-bottom:none;border-left:none;border-rightnone'>&nbsp;</td>                                                                              \n");
                    htmlStr.Append("  <td class=xl11715939 style='border-top:none;border-bottom:none;border-left:none;border-right:1pt solid black'>&nbsp;</td>                                                                                                                 \n");
                    htmlStr.Append(" </tr>                                                                                                                                             \n");



                    htmlStr.Append("	<table  border=0>                                                                                                                                                                                                 \n");
                    htmlStr.Append("		<tr height=5  style='height: 18pt'>                                                                                                                                                                \n");

                    htmlStr.Append("		</tr>      																																														\n");
                    htmlStr.Append("	</table>             																																										\n");

                }


            }// for k                                                                                                                             
            htmlStr.Append(" <tr class=xl8715939 height=22 style='mso-height-source:userset;height:16.95pt'>                                                                   \n");
            htmlStr.Append("  <td height=22 class=xl12615939 style='height:16.95pt;border-top:1pt solid black;'>&nbsp;</td>                                                                                \n");
            htmlStr.Append("  <td class=xl12715939 style='height:16.95pt;border-top:1pt solid black;'>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("  <td class=xl12815939 width=127 style='width:95pt;border-top:1pt solid black;'>&nbsp;</td>                                                                                    \n");
            htmlStr.Append("  <td class=xl12815939 width=62 style='width:47pt;border-top:1pt solid black;'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("  <td class=xl12815939 width=41 style='width:31pt;border-top:1pt solid black;'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("  <td class=xl12815939 width=34 style='width:31.25pt;border-top:1pt solid black;'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("  <td colspan=6 class=xl17515939 width=264 style='width:197pt;border-top:1pt solid black;border-right:1pt solid black'>C&#7897;ng                                                                          \n");
            htmlStr.Append("  ti&#7873;n hàng (<font class='font1015939'>Total net amount</font><font                                                                          \n");
            htmlStr.Append("  class='font815939'>):</font></td>                                                                                                                \n");
            htmlStr.Append("  <td colspan=2 class=xl17615939 style='height:16.95pt;border-top:1pt solid black;'>&nbsp;" + dt.Rows[0]["netamount_display"] + "</td>                                                                                                       \n");
            htmlStr.Append("  <td class=xl12915939 style='height:16.95pt;border-top:1pt solid black;'>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append(" </tr>                                                                                                                                             \n");

            htmlStr.Append(" <tr class=xl8715939 height=22 style='mso-height-source:userset;height:16.95pt'>                                                                   \n");
            htmlStr.Append("  <td height=22 class=xl13015939 style='height:16.95pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                \n");
            htmlStr.Append("  <td colspan=5 class=xl17815939 style='height:16.95pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>Thu&#7871; su&#7845;t GTGT (<font                                                                                 \n");
            htmlStr.Append("  class='font915939'>VAT Rate</font><font class='font815939'>):<span                                                                               \n");
            htmlStr.Append("  style='mso-spacerun:yes'>       </span>" + dt.Rows[0]["TaxRate"] + "</font></td>                                                                                             \n");
            htmlStr.Append("  <td colspan=6 class=xl17915939 width=264 style='width:197pt;border-top:none;border-bottom:none;border-left:none;border-right:1pt solid black;'>Ti&#7873;n                                                                          \n");
            htmlStr.Append("  thu&#7871; GTGT (<font class='font1015939'>VAT amount</font><font                                                                                \n");
            htmlStr.Append("  class='font815939'>):</font></td>                                                                                                                \n");
            htmlStr.Append("  <td colspan=2 class=xl18015939 style='height:16.95pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;" + dt.Rows[0]["vatamount_display"] + "</td>                                                                                                       \n");
            htmlStr.Append("  <td class=xl13115939 style='height:16.95pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append(" </tr>                                                                                                                                             \n");

            htmlStr.Append(" <tr class=xl8715939 height=22 style='mso-height-source:userset;height:16.95pt;border-top:none;border-bottom:1pt solid black;border-left:none;border-right:none'>                                                                   \n");
            htmlStr.Append("  <td height=22 class=xl13015939 style='height:16.95pt;border-top:none;border-bottom:1pt solid black;border-left:none;border-right:none'>&nbsp;</td>                                                                                \n");
            htmlStr.Append("  <td class=xl10515939 style='border-top:none;border-bottom:1pt solid black;border-left:none;border-right:none'>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("  <td class=xl13215939 width=127 style='width:95pt;border-top:none;border-bottom:1pt solid black;border-left:none;border-right:none'>&nbsp;</td>                                                                                    \n");
            htmlStr.Append("  <td class=xl13215939 width=62 style='width:47pt;border-top:none;border-bottom:1pt solid black;border-left:none;border-right:none'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("  <td class=xl13215939 width=41 style='width:31pt;border-top:none;border-bottom:1pt solid black;border-left:none;border-right:none'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("  <td class=xl13215939 width=34 style='width:31.25pt;border-top:none;border-bottom:1pt solid black;border-left:none;border-right:none'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("  <td colspan=6 class=xl17815939 style='border-top:none;border-bottom:1pt solid black;border-left:none;border-right:1pt solid black;'>T&#7893;ng c&#7897;ng ti&#7873;n thành toán (<font                                                                \n");
            htmlStr.Append("  class='font1015939'>Total of payment</font><font class='font815939'>):</font></td>                                                               \n");
            htmlStr.Append("  <td colspan=2 class=xl18015939  style='width:95pt;border-top:none;border-bottom:1pt solid black;border-left:none;border-right:none'>&nbsp;" + dt.Rows[0]["totalamount_display"] + "</td>                                                                                                       \n");
            htmlStr.Append("  <td class=xl13115939  style='width:95pt;border-top:none;border-bottom:1pt solid black;border-left:none;border-right:none'>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append(" </tr>                                                                                                                                             \n");

            htmlStr.Append(" <tr class=xl8715939 height=21 style='height:19.5pt'>                                                                                              \n");
            htmlStr.Append("  <td height=21 class=xl12615939 style='height:19.5pt'>&nbsp;</td>                                                                                 \n");
            htmlStr.Append("  <td colspan=13 class=xl18215939>Số tiền viết bằng                                                                        \n");
            htmlStr.Append("  chữ (<font class='font1015939'>In words</font><font class='font615939'>):<span                                                             \n");
            htmlStr.Append("  style='mso-spacerun:yes;font-style:italic;font-weight:700'>&nbsp;" + read_prive + ". </span></font></td>                                                                                                    \n");
            htmlStr.Append("  <td class=xl13315939 width=12 style='width:9pt'>&nbsp;</td>                                                                                      \n");
            htmlStr.Append(" </tr>                                                                                                                                             \n");
            htmlStr.Append(" <tr class=xl8715939 height=21 style='height:19.5pt'>                                                                                              \n");
            htmlStr.Append("  <td height=21 class=xl13015939 style='height:19.5pt;border-bottom:1pt solid windowtext;border-top:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                 \n");
            htmlStr.Append("  <td class=xl8015939 style='width:95pt;border-bottom:1pt solid windowtext;border-top:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("  <td class=xl13415939 width=127 style='width:95pt;border-bottom:1pt solid windowtext;border-top:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                    \n");
            htmlStr.Append("  <td class=xl13415939 width=62 style='width:47pt;border-bottom:1pt solid windowtext;border-top:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("  <td class=xl13415939 width=41 style='width:31pt;border-bottom:1pt solid windowtext;border-top:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("  <td class=xl13415939 width=34 style='width:31.25pt;border-bottom:1pt solid windowtext;border-top:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("  <td class=xl13415939 width=77 style='width:72.5‬pt;border-bottom:1pt solid windowtext;border-top:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("  <td class=xl13415939 width=78 style='width:72.5‬pt;border-bottom:1pt solid windowtext;border-top:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("  <td class=xl13415939 width=6 style='width:5pt;border-bottom:1pt solid windowtext;border-top:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                       \n");
            htmlStr.Append("  <td class=xl13415939 width=49 style='width:46.25‬pt;border-bottom:1pt solid windowtext;border-top:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("  <td class=xl13415939 width=48 style='width:45pt;border-bottom:1pt solid windowtext;border-top:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("  <td class=xl13415939 width=6 style='width:5pt;border-bottom:1pt solid windowtext;border-top:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                       \n");
            htmlStr.Append("  <td class=xl13415939 width=45 style='width:43.75pt;border-bottom:1pt solid windowtext;border-top:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("  <td class=xl13415939 width=85 style='width:81.31.25pt;border-bottom:1pt solid windowtext;border-top:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("  <td class=xl13515939 width=12 style='width:9pt;border-bottom:1pt solid windowtext;border-top:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                      \n");
            htmlStr.Append(" </tr>                                                                                                                                             \n");

            htmlStr.Append(" <tr class=xl9120842 height = 21                                                                                                                                                         \n");
            htmlStr.Append(" style='mso-height-source: userset; height: 15.75pt'>                                                                                                                                    \n");
            htmlStr.Append("     <td height = 21 class=xl14120842 style = 'height: 15.75pt' > &nbsp;</td>                                                                                                            \n");
            htmlStr.Append("     <td colspan = 4 class=xl18320842 width = 263 style='width: 198pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>Ng&#432;&#7901;i                                                                                                  \n");
            htmlStr.Append(" 	    th&#7921;c hi&#7879;n chuy&#7875;n &#273;&#7893;i (Converter)</td>                                                                                                               \n");
            htmlStr.Append("     <td colspan = 4 class=xl18420842 width = 195 style='width: 145pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>Ng&#432;&#7901;i                                                                                                  \n");
            htmlStr.Append(" 	    mua hàng(Buyer)</td>                                                                                                                                                             \n");
            htmlStr.Append("     <td colspan = 5 class=xl17620842 width = 233 style='width: 175pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>Ng&#432;&#7901;i                                                                                                  \n");
            htmlStr.Append(" 	    bán hàng(Seller)</td>                                                                                                                                                            \n");
            htmlStr.Append("     <td class=xl14020842 width = 12 style='width: 9pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                                                      \n");
            htmlStr.Append(" </tr>                                                                                                                                                                                   \n");
            htmlStr.Append(" <tr class=xl8120842 height = 18                                                                                                                                                         \n");
            htmlStr.Append("     style='mso-height-source: userset; height: 13.95pt'>                                                                                                                                \n");
            htmlStr.Append("     <td height = 18 class=xl14220842 style = 'height: 13.95pt' > &nbsp;</td>                                                                                                            \n");
            htmlStr.Append("     <td colspan = 3 class=xl18520842 width = 222 style='width: 167pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>(Ký,                                                                                                              \n");
            htmlStr.Append(" 	    ghi rõ h&#7885; tên)</td>                                                                                                                                                        \n");
            htmlStr.Append("     <td class=xl14320842 width = 41 style='width: 31pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                                                     \n");
            htmlStr.Append("     <td colspan = 4 class=xl18520842 width = 195 style='width: 145pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>(Ký,                                                                                                              \n");
            htmlStr.Append(" 	    ghi rõ h&#7885; tên)</td>                                                                                                                                                        \n");
            htmlStr.Append("     <td colspan = 5 class=xl14320842 width = 233 style='width: 175pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>(Ký,                                                                                                              \n");
            htmlStr.Append(" 	    &#273;óng d&#7845;u, ghi rõ h&#7885; tên)</td>                                                                                                                                   \n");
            htmlStr.Append("     <td class=xl14420842 width = 12 style='width: 9pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                                                      \n");
            htmlStr.Append(" </tr>                                                                                                                                                                                   \n");
            htmlStr.Append(" <tr class=xl14820842 height = 20                                                                                                                                                        \n");
            htmlStr.Append("     style='mso-height-source: userset; height: 15.45pt'>                                                                                                                                \n");
            htmlStr.Append("     <td height = 20 class=xl14520842 style = 'height: 15.45pt;border-top:none;border-bottom:none;border-left:none;border-right:none' > &nbsp;</td>                                                                                                            \n");
            htmlStr.Append("     <td colspan = 3 class=xl18620842 width = 222 style='width: 167pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>(Signature                                                                                                        \n");
            htmlStr.Append(" 	    &amp; full name)</td>                                                                                                                                                            \n");
            htmlStr.Append("     <td class=xl14620842 width = 41 style='width: 31pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                                                     \n");
            htmlStr.Append("     <td colspan = 4 class=xl18620842 width = 195 style='width: 145pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>(Signature                                                                                                        \n");
            htmlStr.Append(" 	    &amp; full name)</td>                                                                                                                                                            \n");
            htmlStr.Append("     <td colspan = 5 class=xl14620842 width = 233 style='width: 175pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>(Signature,                                                                                                       \n");
            htmlStr.Append(" 	    stamp &amp; full name)</td>                                                                                                                                                      \n");
            htmlStr.Append("     <td class=xl14720842 width = 12 style='width: 9pt'>&nbsp;</td>                                                                                                                      \n");
            htmlStr.Append(" </tr>                                                                                                                                                                                   \n");

            htmlStr.Append("<tr height = 13 style = 'mso-height-source:userset;height:15.05pt' >                                                                                                               \n");
            htmlStr.Append("  <td height = 13 class=xl7720842 style ='height:15.05pt;border-top:none;border-bottom:none;border-right:none' > &nbsp;</td>                                                     \n");
            htmlStr.Append("  <td class=xl7320842 style = 'width:95pt;border-top:none;border-bottom:none;border-right:none;border-left:none' > &nbsp;</td>                                                     \n");
            htmlStr.Append("  <td class=xl8320842 width = 127 style='width:95pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                             \n");
            htmlStr.Append("  <td class=xl8320842 width = 62 style='width:47pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                              \n");
            htmlStr.Append("  <td class=xl8320842 width = 41 style='width:31pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                              \n");
            htmlStr.Append("  <td class=xl8320842 width = 34 style='width:25pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                              \n");
            htmlStr.Append("  <td class=xl8320842 width = 77 style='width:58pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                              \n");
            htmlStr.Append("  <td class=xl8320842 width = 78 style='width:58pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                              \n");
            htmlStr.Append("  <td class=xl8320842 width = 6 style='width:4pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                \n");
            htmlStr.Append("  <td class=xl8320842 width = 49 style='width:37pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                              \n");
            htmlStr.Append("  <td class=xl8320842 width = 48 style='width:36pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                              \n");
            htmlStr.Append("  <td class=xl8320842 width = 6 style='width:4pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                \n");
            htmlStr.Append("  <td class=xl8320842 width = 45 style='width:34pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                              \n");
            htmlStr.Append("  <td class=xl8320842 width = 85 style='width:64pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                              \n");
            htmlStr.Append("  <td class=xl14920842 width = 12 style='width:9pt;border-top:none;border-bottom:none;border-left:none;'>&nbsp;</td>                                                               \n");
            htmlStr.Append("</tr>                                                                                                                                                                              \n");
            htmlStr.Append(" <tr height=22 style='height:21‬pt'>                                                                                                              \n");
            htmlStr.Append("  <td height=22 class=xl8415939 style='height:21‬pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                  \n");
            htmlStr.Append("  <td class=xl7315939 style='border-top:none;border-bottom:none;border-left:none;border-right:none'></td>                                                                                                                  \n");
            htmlStr.Append("  <td class=xl8215939 width=127 style='width:95pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>" + dt.Rows[0]["convert_name"] + "</td>                                                                                     \n");
            htmlStr.Append("  <td class=xl8215939 width=62 style='width:47pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                      \n");
            htmlStr.Append("  <td class=xl8215939 width=41 style='width:31pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                      \n");
            htmlStr.Append("  <td class=xl8215939 width=34 style='width:31.25pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                      \n");
            htmlStr.Append("  <td class=xl8215939 width=77 style='width:72.5‬pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                      \n");
            htmlStr.Append("  <td class=xl7315939 style='border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                                                   \n");
            htmlStr.Append("  <td class=xl6515939 colspan=3 style='border-bottom:none;border-left:1pt solid windowtext;border-right:none'>Signature Valid</td>                                                                                               \n");

            if (dt.Rows[0]["sign_yn"].ToString() == "Y")
            {
                htmlStr.Append("  <td align=left class=xl6715939 valign=top style='border-bottom:none;border-left:none;border-right:none' ><![if !vml]><span style='mso-ignore:vglayout;                                                                          \n");
                htmlStr.Append("  position:absolute;z-index:1;margin-left:-1px;margin-top:7px;width:81px;                                                                           \n");
                htmlStr.Append("  height:58px'><img width=81 height=58                                                                                                             \n");
                htmlStr.Append("  src='D:/webproject/e-invoice-ws/02.Web/EInvoice/img/check_signed.png'                               \n");
                htmlStr.Append("  v:shapes='Picture_x0020_8'></span><![endif]><span style='mso-ignore:vglayout2'>                                                                  \n");
                htmlStr.Append("  <table cellpadding=0 cellspacing=0>                                                                                                              \n");
                htmlStr.Append("   <tr>                                                                                                                                            \n");
                htmlStr.Append("    <td height=22  width=6 style='height:21‬pt;width:5pt;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                              \n");
                htmlStr.Append("   </tr>                                                                                                                                           \n");
                htmlStr.Append("  </table>                                                                                                                                         \n");
                htmlStr.Append("  </span></td>        \n");
            }
            else
            {
                htmlStr.Append("  <td align=left class=xl6715939 valign=top style='border-bottom:none;border-left:none;border-right:none' ></td>        \n");
            }
            htmlStr.Append("  <td class=xl6715939 style='border-top:1pt solid windowtext;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("  <td class=xl6815939 style='border-top:1pt solid windowtext;border-bottom:none;border-left:none;border-right:1pt solid windowtext'>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("  <td class=xl14915939 width=12 style='width:9pt'>&nbsp;</td>                                                                                      \n");
            htmlStr.Append(" </tr>                                                                                                                                             \n");

            htmlStr.Append(" <tr height=21 style='height:19.5pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>                                                                                                              \n");
            htmlStr.Append("  <td height=21 class=xl8415939 style='height:19.5pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                  \n");
            htmlStr.Append("  <td class=xl15215939 style='border-top:none;border-bottom:none;border-left:none;border-right:none'>" + dt.Rows[0]["convert_date"] + "</td>                                                                                                                     \n");
            htmlStr.Append("  <td class=xl8215939 width=127 style='width:95pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("  <td class=xl8215939 width=62 style='width:47pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                      \n");
            htmlStr.Append("  <td class=xl8215939 width=41 style='width:31pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                      \n");
            htmlStr.Append("  <td class=xl8215939 width=34 style='width:31.25pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                      \n");
            htmlStr.Append("  <td class=xl8215939 width=77 style='width:72.5‬pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                      \n");
            htmlStr.Append("  <td class=xl7315939 style='border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("  <td colspan=6 class=xl18615939 style='border-right:1pt solid black;border-top:none;border-bottom:none;border-left:1pt solid black;'><font                                                                       \n");
            htmlStr.Append("  class='font1315939'>Được ký bởi:</font><font                                                                               \n");
            htmlStr.Append("  class='font1215939'> </font><font class='font1415939' style='font-size:9pt'>" + dt.Rows[0]["SignedBy"] + "</font>                                                                       \n");
            htmlStr.Append("  <td class=xl14915939 width=12 style='width:9pt'>&nbsp;</td>                                                                                      \n");
            htmlStr.Append(" </tr>                                                                                                                                             \n");
            htmlStr.Append(" <tr height=21 style='height:19.5pt'>                                                                                                              \n");
            htmlStr.Append("  <td height=21 class=xl8415939 style='height:19.5pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                  \n");
            htmlStr.Append("  <td class=xl15015939 colspan=2 style='border-top:none;border-bottom:none;border-left:none;border-right:none'>Mã CQT: " + dt.Rows[0]["cqt_mccqt_id"] + "</td>                                                                             \n");
            htmlStr.Append("  <td class=xl8215939 width=62 style='width:47pt;border-top:none;border-bottom:none;border-left:none;border-right:none'></td>                                                                                      \n");
            htmlStr.Append("  <td class=xl8215939 width=41 style='width:31pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                      \n");
            htmlStr.Append("  <td class=xl8215939 width=34 style='width:31.25pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                      \n");
            htmlStr.Append("  <td class=xl8215939 width=77 style='width:72.5‬pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                      \n");
            htmlStr.Append("  <td class=xl7315939 style='border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                                                   \n");
            htmlStr.Append("  <td class=xl6915939 colspan=5 style='width:47pt;border-top:none;border-bottom:1pt solid black;border-left:1pt solid black;border-right:none'>Ngày Ký: <font class='font1515939'>" + dt.Rows[0]["SignedDate"] + "</font></td>                                                          \n");
            htmlStr.Append("  <td class=xl7115939 style='width:47pt;border-top:none;border-bottom:1pt solid black;border-left:none;border-right:1pt solid black'>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("  <td class=xl14915939 width=12 style='width:9pt'>&nbsp;</td>                                                                                      \n");
            htmlStr.Append(" </tr>                                                                                                                                             \n");
            htmlStr.Append(" <tr height=21 style='height:19.5pt'>                                                                                                              \n");
            htmlStr.Append("  <td height=21 class=xl8415939 style='height:19.5pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                  \n");
            htmlStr.Append("  <td class=xl8015939 colspan=2 style='border-top:none;border-bottom:none;border-left:none;border-right:none'>Tra cứu tại Website: <font class='font1226793'>" + dt.Rows[0]["tracuuwebsite"] + "</font></td>                                                                              \n");
            htmlStr.Append("  <td class=xl8215939 width=62 style='width:47pt;border-top:none;border-bottom:none;border-left:none;border-right:none'></td>                                                                                      \n");
            htmlStr.Append("  <td class=xl8215939 width=41 style='width:31pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                      \n");
            htmlStr.Append("  <td class=xl8215939 width=34 style='width:31.25pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                      \n");
            htmlStr.Append("  <td class=xl8215939 width=77 style='width:72.5‬pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                      \n");
            htmlStr.Append("  <td class=xl8215939 width=78 style='width:72.5‬pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                      \n");
            htmlStr.Append("  <td class=xl8215939 width=49 style='width:46.25‬pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>Mã nh&#7853;n hóa &#273;&#417;n:  <font class='font1226793'>" + dt.Rows[0]["matracuu"] + "&nbsp;</font></td>                                                                                      \n");

            htmlStr.Append("  <td class=xl8215939 width=6 style='width:5pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                        \n");
            htmlStr.Append("  <td class=xl8215939 width=48 style='width:45pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                      \n");
            htmlStr.Append("  <td class=xl8215939 width=6 style='width:5pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                        \n");
            htmlStr.Append("  <td class=xl8215939 width=45 style='width:43.75pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                      \n");
            htmlStr.Append("  <td class=xl8215939 width=85 style='width:81.31.25pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                      \n");
            htmlStr.Append("  <td class=xl14915939 width=12 style='width:9pt;border-top:none;border-bottom:none;border-left:none;border-right:none'>&nbsp;</td>                                                                                      \n");
            htmlStr.Append(" </tr>                                                                                                                                             \n");
            htmlStr.Append(" <tr height=21 style='height:19.5pt'>                                                                                                              \n");
            htmlStr.Append("  <td height=21 class=xl9615939 style='height:19.5pt'>&nbsp;</td>                                                                                  \n");
            htmlStr.Append("  <td colspan=13 class=xl18415939>(Cần kiểm tra, đối                                                                        \n");
            htmlStr.Append("  chiếu khi lập, giao nhận hóa đơn)</td>												                                \n");
            htmlStr.Append("  <td class=xl15115939 width=12 style='width:9pt'>&nbsp;</td>                                                                                      \n");
            htmlStr.Append(" </tr>                                                                                                                                             \n");

            htmlStr.Append(" <![if supportMisalignedColumns]>                                                                                                                  \n");
            htmlStr.Append(" <tr height=0 style='display:none'>                                                                                                                \n");
            htmlStr.Append("  <td width=6 style='width:5pt'></td>                                                                                                              \n");
            htmlStr.Append("  <td width=33 style='width:31.25pt'></td>                                                                                                            \n");
            htmlStr.Append("  <td width=127 style='width:95pt'></td>                                                                                                           \n");
            htmlStr.Append("  <td width=62 style='width:47pt'></td>                                                                                                            \n");
            htmlStr.Append("  <td width=41 style='width:31pt'></td>                                                                                                            \n");
            htmlStr.Append("  <td width=34 style='width:31.25pt'></td>                                                                                                            \n");
            htmlStr.Append("  <td width=77 style='width:72.5‬pt'></td>                                                                                                            \n");
            htmlStr.Append("  <td width=78 style='width:72.5‬pt'></td>                                                                                                            \n");
            htmlStr.Append("  <td width=6 style='width:5pt'></td>                                                                                                              \n");
            htmlStr.Append("  <td width=49 style='width:46.25‬pt'></td>                                                                                                            \n");
            htmlStr.Append("  <td width=48 style='width:45pt'></td>                                                                                                            \n");
            htmlStr.Append("  <td width=6 style='width:5pt'></td>                                                                                                              \n");
            htmlStr.Append("  <td width=45 style='width:43.75pt'></td>                                                                                                            \n");
            htmlStr.Append("  <td width=85 style='width:81.31.25pt'></td>                                                                                                            \n");
            htmlStr.Append("  <td width=12 style='width:9pt'></td>                                                                                                             \n");
            htmlStr.Append(" </tr>                                                                                                                                             \n");
            htmlStr.Append(" <![endif]>                                                                                                                                        \n");
            htmlStr.Append("</table>                                                                                                                                           \n");
            htmlStr.Append("<table>                                                                                                                                           \n");
            htmlStr.Append(" <tr height=18 style='height:16.5‬pt;'>                                                                                                              \n");
            htmlStr.Append("  <td height=18 class=xl7315939 style='height:16.5‬pt'>&nbsp;</td>                                                                                  \n");
            htmlStr.Append("  <td colspan=13 class=xl18515939>" + dt.Rows[0]["CONTRACT_INFO_EI"] + "</td>                                                                                                         \n");
            htmlStr.Append("  <td class=xl7215939 style='border-top:none'>&nbsp;</td>                                                                                          \n");
            htmlStr.Append(" </tr>                                                                                                                                             \n");
            htmlStr.Append("</table>                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                   \n");
            htmlStr.Append("</div>  \n");
            htmlStr.Append("</body>                                                                                                                                                                                                 \n");
            htmlStr.Append("</html>               \n");

            string filePath = "";

           

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