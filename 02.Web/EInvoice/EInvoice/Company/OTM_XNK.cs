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

    public class OTM_XNK
    {
        public static string View(string tei_einvoice_m_pk, string tei_company_pk, string dbName)
        {
            /*string dbUser = "genuwin", dbPwd = "genuwin2";//NOBLANDBD  EINVOICE_252
            string _conString = "Data Source={0};User Id={1};Password={2};Unicode=true";
            _conString = String.Format(_conString, dbName, dbUser, dbPwd);*/
            string _conString = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=123.30.104.243)(PORT=1941))(CONNECT_DATA=(SERVER=dedicated)(SERVICE_NAME=NOBLANDBD)));User ID=genuwin;Password=genuwin2";


            string Procedure = "stacfdstac71_r_02_view_einv_v2"; //stacfdstac71_r_02_xk
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
            int[] page_index = new int[10] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            int v_index = -1, rowsPerPage = 20;

            int v_countNumberOfPages = 0;
            int total_countLenght = 0;
            int count_col = 0;
            string l_finish = "N";
            int count_col_index = 0;
            for (int i = 0; i < 99; i++)
            {
                count_col_index = 0;
                total_countLenght = 0;
                for (int j = count_col; j < v_count; j++)
                {

                    int count_row = countLength(dt_d.Rows[j]["ITEM_NAME"].ToString());
                    if (count_row > 0)
                    {
                        total_countLenght += count_row;

                    }
                    else
                    {
                        total_countLenght += 1;
                    }

                    if (count_col == v_count - 1)
                    {
                        if (total_countLenght > pos)
                        {
                            page[i] = count_col_index;
                            page_index[i] = total_countLenght - 1;
                            page[i + 1] = 1;
                            page_index[i + 1] = 1;
                            l_finish = "Y";
                            count_col++;
                            count_col_index++;
                            break;
                        }
                        else
                        {
                            page[i] = count_col_index + 1;
                            page_index[i] = total_countLenght - 1;
                            l_finish = "Y";
                            count_col++;
                            count_col_index++;
                            break;
                        }


                    }
                    else if (total_countLenght > pos && total_countLenght < pos_lv)
                    {
                        if (pos_lv - total_countLenght < 3)
                        {
                            page[i] = count_col_index + 1;
                            page_index[i] = total_countLenght - 1;
                            count_col++;
                            count_col_index++;
                            break;//continue;	
                        }
                    }
                    else if (total_countLenght == pos_lv)
                    {
                        page[i] = count_col_index;
                        page_index[i] = total_countLenght + 1;
                        count_col++;
                        count_col_index++;
                        break;//continue;
                    }
                    count_col++;
                    count_col_index++;
                }
                if (l_finish == "Y")
                {
                    break;
                }

            }

            for (int i = 0; i < 10; i++)
            {
                if (page[i] > 0)
                {
                    v_countNumberOfPages++;
                }
            }

            string[] parts = dt.Rows[0]["SELLER_ADDRESS"].ToString().Split(',');
            string l_add = parts[0] + "," + parts[1] + "," + parts[2] + ",";
            string l_add1 = parts[3] + "," + parts[4] + "," + parts[5];// + "," + parts[6];

            string[] parts1 = dt.Rows[0]["SELLER_ADDRESS1"].ToString().Split(',');
            string l_adden = parts1[0] + "," + parts1[1] + "," + parts1[2] + ",";
            string l_add1en = parts1[3] + "," + parts1[4] + "," + parts1[5];// + "," + parts1[6];

            /* string read_prive = "", read_en = "", read_amount = "", amount_vat = "", amount_total = "", amount_trans = "", amount_net = "", lb_amount_trans = "";


             if (dt.Rows[0]["CurrencyCodeUSD"].ToString() == "VND")
             {
                 lb_amount_trans = "";
                 amount_trans = "";
                 amount_total = dt.Rows[0]["TOT_AMT_BK_93"].ToString();
                 amount_vat = dt.Rows[0]["VAT_BK_AMT_92"].ToString();
                 amount_net = dt.Rows[0]["NET_BK_AMT_90"].ToString();
                 read_prive = NumberToTextVN(Decimal.Parse(dt.Rows[0]["amt_tr_read_95"].ToString()));
             }
             else
             {
                 lb_amount_trans = dt.Rows[0]["EXCHANGERATE"].ToString();
                 amount_trans = dt.Rows[0]["TOT_AMT_BK_93"].ToString();
                 amount_total = dt.Rows[0]["tot_amt_tr_94"].ToString();
                 amount_vat = dt.Rows[0]["VAT_TR_AMT_DIS_TR_91"].ToString();
                 amount_net = dt.Rows[0]["NET_TR_AMT_DIS_TR_89"].ToString();

                 read_prive = Num2VNText(dt.Rows[0]["amt_tr_read_95"].ToString(), dt.Rows[0]["CurrencyCodeUSD"].ToString());
             }
             read_prive = read_prive.Substring(0, 2) + read_prive.Substring(2, read_prive.Length - 2).ToLower() + '.';
          */
            //read_en = dt.Rows[0]["TotalAmountInWord"].ToString();
            int end = 0;
            int count = count_page_v + r;
            double height = 130;
            StringBuilder htmlStr = new StringBuilder("");
            string heigh = "", heigh_d = "";


            htmlStr.Append("<!DOCTYPE html PUBLIC '-//W3C//DTD HTML 4.01 Transitional//EN' 'http://www.w3.org/TR/html4/loose.dtd'>																									\n");
            htmlStr.Append("<html>                                                                                                                                                                                                  \n");
            htmlStr.Append("<head>                                                                                                                                                                                                  \n");
            htmlStr.Append("<meta http-equiv='Content-Type' content='text/html; charset=UTF-8'>                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append("<script type='text/javascript'                                                                                                                                                                          \n");
            htmlStr.Append("	src='${pageContext.request.contextPath}/system/syscommand.js'></script>                                                                                                                             \n");
            htmlStr.Append("<title>Report E-Invoice</title>                                                                                                                                                                         \n");
            //htmlStr.Append("<!-- Normalize or reset CSS with your favorite library -->                                                                                                                                              \n");
            //htmlStr.Append("<link rel='stylesheet'                                                                                                                                                                                  \n");
            //htmlStr.Append("	href='https://cdnjs.cloudflare.com/ajax/libs/normalize/3.0.3/normalize.css'>                                                                                                                        \n");
            //htmlStr.Append("                                                                                                                                                                                                        \n");
            //htmlStr.Append("<!-- Load paper.css for happy printing -->                                                                                                                                                              \n");
            ///htmlStr.Append("<link rel='stylesheet'                                                                                                                                                                                  \n");
            //htmlStr.Append("	href='https://cdnjs.cloudflare.com/ajax/libs/paper-css/0.2.3/paper.css'>                                                                                                                            \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append("<!-- Set page size here: A5, A4 or A3 -->                                                                                                                                                               \n");
            htmlStr.Append("<!-- Set also 'landscape' if you need -->                                                                                                                                                               \n");
            htmlStr.Append("<style>                                                                                                                                                                                                 \n");
            htmlStr.Append("@page {                                                                                                                                                                                                 \n");
            htmlStr.Append("	size: A4                                                                                                                                                                                            \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("</style>                                                                                                                                                                                                \n");
            ///htmlStr.Append("<link href='https://fonts.googleapis.com/css?family=Tangerine:700'                                                                                                                                      \n");
            //htmlStr.Append("	rel='stylesheet' type='text/css'>                                                                                                                                                                   \n");
            htmlStr.Append("<style>                                                                                                                                                                                                 \n");
            htmlStr.Append("/*body   { font-family: serif }                                                                                                                                                                         \n");
            htmlStr.Append("    h1     { font-family: 'Tangerine', cursive; font-size: 40pt; line-height: 18mm}                                                                                                                     \n");
            htmlStr.Append("    h2, h3 { font-family: 'Tangerine', cursive; font-size: 24pt; line-height: 7mm }                                                                                                                     \n");
            htmlStr.Append("    h4     { font-size: 16.25pt; line-height: 1mm }                                                                                                                                                        \n");
            htmlStr.Append("    h2 + p { font-size: 18pt; line-height: 7mm }                                                                                                                                                        \n");
            htmlStr.Append("    h3 + p { font-size: 14pt; line-height: 7mm }                                                                                                                                                        \n");
            htmlStr.Append("    li     { font-size: 11pt; line-height: 5mm }                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append("    h1      { margin: 0 }                                                                                                                                                                               \n");
            htmlStr.Append("    h1 + ul { margin: 2mm 0 5mm }                                                                                                                                                                       \n");
            htmlStr.Append("    h2, h3  { margin: 0 3mm 3mm 0; float: left }                                                                                                                                                        \n");
            htmlStr.Append("    h2 + p,                                                                                                                                                                                             \n");
            htmlStr.Append("    h3 + p  { margin: 0 0 3mm 50mm }                                                                                                                                                                    \n");
            htmlStr.Append("    //h4      { margin: 1mm 0 0 2mm; border-bottom: 1px solid black }                                                                                                                                   \n");
            htmlStr.Append("    h4 + ul { margin: 5mm 0 0 50mm }                                                                                                                                                                    \n");
            htmlStr.Append("    article { border: 4px double black; padding: 5mm 10mm; border-radius: 3mm }*/                                                                                                                       \n");
            htmlStr.Append("body {                                                                                                                                                                                                  \n");
            //htmlStr.Append("	color: blue;                                                                                                                                                                                        \n");
            htmlStr.Append("	font-size: 100%;                                                                                                                                                                                    \n");
            htmlStr.Append("	background-image: url('assets/Solution.jpg');                                                                                                                                                       \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append("h1 {                                                                                                                                                                                                    \n");
            htmlStr.Append("	color: #00FF00;                                                                                                                                                                                     \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append("p {                                                                                                                                                                                                     \n");
            htmlStr.Append("	color: rgb(0, 0, 255)                                                                                                                                                                               \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append("headline1 {                                                                                                                                                                                             \n");
            htmlStr.Append("	background-image: url(assets/Solution.jpg);                                                                                                                                                         \n");
            htmlStr.Append("	background-repeat: no-repeat;                                                                                                                                                                       \n");
            htmlStr.Append("	background-position: left top;                                                                                                                                                                      \n");
            htmlStr.Append("	padding-top: 68px;                                                                                                                                                                                  \n");
            htmlStr.Append("	margin-bottom: 50px;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append("headline2 {                                                                                                                                                                                             \n");
            htmlStr.Append("	background-image: url(images/newsletter_headline2.gif);                                                                                                                                             \n");
            htmlStr.Append("	background-repeat: no-repeat;                                                                                                                                                                       \n");
            htmlStr.Append("	background-position: left top;                                                                                                                                                                      \n");
            htmlStr.Append("	padding-top: 68px;                                                                                                                                                                                  \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append("	<!--																																					 \n");
            htmlStr.Append("	table {                                                                                                                                                  \n");
            htmlStr.Append("		mso-displayed-decimal-separator: '\\.';                                                                                                               \n");
            htmlStr.Append("		mso-displayed-thousand-separator: '\\, ';                                                                                                              \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.font530999 {                                                                                                                                            \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 13.75pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.font65513 {                                                                                                                                             \n");
            htmlStr.Append("		color: windowtext;                                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.font630999 {                                                                                                                                            \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 14.0pt;;                                                                                                                                    \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.font730999 {                                                                                                                                            \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 14.0pt;;                                                                                                                                    \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.font830999 {                                                                                                                                            \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.font930999 {                                                                                                                                            \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.font1030999 {                                                                                                                                           \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 13.75pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.font185513 {                                                                                                                                            \n");
            htmlStr.Append("		color: #0066CC;                                                                                                                                      \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.font1130999 {                                                                                                                                           \n");
            htmlStr.Append("		color: red;                                                                                                                                          \n");
            htmlStr.Append("		font-size: 16.25pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.font1230999 {                                                                                                                                           \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.font1330999 {                                                                                                                                           \n");
            htmlStr.Append("		color: red;                                                                                                                                          \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.font1430999 {                                                                                                                                           \n");
            htmlStr.Append("		color: red;                                                                                                                                          \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.font1530999 {                                                                                                                                           \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 13.75pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.font1630999 {                                                                                                                                           \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl6330999 {                                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 1.0pt;                                                                                                                                    \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: 1.0pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("		border-right: none;                                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                                 \n");
            htmlStr.Append("		border-left: 1.0pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl6430999 {                                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 1.0pt;                                                                                                                                    \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: 1.0pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("		border-right: none;                                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl6530999 {                                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 1.0pt;                                                                                                                                    \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: 1.0pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("		border-right: 1.0pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl6630999 {                                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 13.75pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl6730999 {                                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                                 \n");
            htmlStr.Append("		border-left: 1.0pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl67309991 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 1.0pt;                                                                                                                                    \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                                 \n");
            htmlStr.Append("		border-left: 1.0pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl6830999 {                                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 13.75pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                                    \n");
            htmlStr.Append("		border-right: 1.0pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl68309991 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 1.0pt;                                                                                                                                    \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                                    \n");
            htmlStr.Append("		border-right: 1.0pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl6930999 {                                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 14.0pt;                                                                                                                                    \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl7030999 {                                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl7130999 {                                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 13.75pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                                 \n");
            htmlStr.Append("		border-left: 1.0pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl7230999 {                                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 13.75pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                                    \n");
            htmlStr.Append("		border-right: 1.0pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl7330999 {                                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 13.75pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl7430999 {                                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl7530999 {                                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl7630999 {                                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                                    \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl7730999 {                                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                                 \n");
            htmlStr.Append("		border-left: 1.0pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl7830999 {                                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                                    \n");
            htmlStr.Append("		border-right: 1.0pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl7930999 {                                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl8030999 {                                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 11.88pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: .75pt dotted windowtext;                                                                                                                  \n");
            htmlStr.Append("		border-right: 1.0pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("		border-bottom: .75pt dotted windowtext;                                                                                                               \n");
            htmlStr.Append("		border-left: 1.0pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl8130999 {                                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 11.88pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: '\\#\\,\\#\\#0_\\)\\;\\\\\\(\\#\\,\\#\\#0\\\\\\)';                                                                                                \n");
            htmlStr.Append("		text-align: right;                                                                                                                                   \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: .75pt dotted windowtext;                                                                                                                  \n");
            htmlStr.Append("		border-right: 1.0pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("		border-bottom: .75pt dotted windowtext;                                                                                                               \n");
            htmlStr.Append("		border-left: 1.0pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl8230999 {                                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 14.0pt;;                                                                                                                                    \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: '\\@';                                                                                                                             \n");
            htmlStr.Append("		text-align: right;                                                                                                                                   \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: .75pt dotted windowtext;                                                                                                                  \n");
            htmlStr.Append("		border-right: 1.0pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("		border-bottom: .75pt dotted windowtext;                                                                                                               \n");
            htmlStr.Append("		border-left: none;                                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl8330999 {                                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 13.75pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: .75pt dotted windowtext;                                                                                                                  \n");
            htmlStr.Append("		border-right: 1.0pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("		border-bottom: .75pt dotted windowtext;                                                                                                               \n");
            htmlStr.Append("		border-left: none;                                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl8430999 {                                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                                 \n");
            htmlStr.Append("		border-left: 1.0pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl8530999 {                                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                                    \n");
            htmlStr.Append("		border-right: 1.0pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl8630999 {                                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 14.0pt;;                                                                                                                                    \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                                 \n");
            htmlStr.Append("		border-left: 1.0pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl8730999 {                                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 14.0pt;;                                                                                                                                    \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                                    \n");
            htmlStr.Append("		border-right: 1.0pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl8830999 {                                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: red;                                                                                                                                          \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                                    \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl8930999 {                                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: red;                                                                                                                                          \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                                    \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl9030999 {                                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 13.75pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1.0pt solid windowtext;                                                                                                                \n");
            htmlStr.Append("		border-left: 1.0pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl9130999 {                                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 13.75pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1.0pt solid windowtext;                                                                                                                \n");
            htmlStr.Append("		border-left: none;                                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl9230999 {                                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 13.75pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                                    \n");
            htmlStr.Append("		border-right: 1.0pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("		border-bottom: 1.0pt solid windowtext;                                                                                                                \n");
            htmlStr.Append("		border-left: none;                                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl9330999 {                                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 13.75pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                                 \n");
            htmlStr.Append("		border-left: 1.0pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl9430999 {                                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 13.75pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl9530999 {                                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 14.0pt;;                                                                                                                                    \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl9630999 {                                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 13.75pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                                    \n");
            htmlStr.Append("		border-right: 1.0pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl9730999 {                                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 8.0pt;                                                                                                                                    \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                                    \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl9830999 {                                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 8.0pt;                                                                                                                                    \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl9930999 {                                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: red;                                                                                                                                          \n");
            htmlStr.Append("		font-size: 17.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl10030999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: red;                                                                                                                                          \n");
            htmlStr.Append("		font-size: 16.75pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl10130999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: red;                                                                                                                                          \n");
            htmlStr.Append("		font-size: 17.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl10230999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: red;                                                                                                                                          \n");
            htmlStr.Append("		font-size: 17.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl10330999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 11.88pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: .75pt dotted windowtext;                                                                                                                  \n");
            htmlStr.Append("		border-right: 1.0pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("		border-bottom: .75pt dotted windowtext;                                                                                                               \n");
            htmlStr.Append("		border-left: 1.0pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl10430999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 11.88pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: '\\@';                                                                                                                             \n");
            htmlStr.Append("		text-align: right;                                                                                                                                   \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: .75pt dotted windowtext;                                                                                                                  \n");
            htmlStr.Append("		border-right: none;                                                                                                                                  \n");
            htmlStr.Append("		border-bottom: .75pt dotted windowtext;                                                                                                               \n");
            htmlStr.Append("		border-left: 1.0pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl10530999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: '\\@';                                                                                                                             \n");
            htmlStr.Append("		text-align: right;                                                                                                                                   \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: 1.0pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("		border-right: none;                                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1.0pt solid windowtext;                                                                                                                \n");
            htmlStr.Append("		border-left: none;                                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl10630999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border: 1.0pt solid windowtext;                                                                                                                       \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl10730999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 13.75pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl10830999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border: 1.0pt solid windowtext;                                                                                                                       \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl10930999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: 1.0pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("		border-right: none;                                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                                 \n");
            htmlStr.Append("		border-left: 1.0pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl11030999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: 1.0pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("		border-right: none;                                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl11130999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1.0pt solid windowtext;                                                                                                                \n");
            htmlStr.Append("		border-left: 1.0pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl11230999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1.0pt solid windowtext;                                                                                                                \n");
            htmlStr.Append("		border-left: none;                                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl11330999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 11.88pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                                    \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: .75pt dotted windowtext;                                                                                                                  \n");
            htmlStr.Append("		border-right: none;                                                                                                                                  \n");
            htmlStr.Append("		border-bottom: .75pt dotted windowtext;                                                                                                               \n");
            htmlStr.Append("		border-left: 1.0pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl113309991 {                                                                                                                                           \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 21.25pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: 1.0pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("		border-right: none;                                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1.0pt solid windowtext;                                                                                                                \n");
            htmlStr.Append("		border-left: 1.0pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("		color: red;                                                                                                                                          \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl11430999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: 1.0pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("		border-right: none;                                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1.0pt solid windowtext;                                                                                                                \n");
            htmlStr.Append("		border-left: none;                                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl11530999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: 1.0pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("		border-right: 1.0pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("		border-bottom: 1.0pt solid windowtext;                                                                                                                \n");
            htmlStr.Append("		border-left: none;                                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl11630999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl11730999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 13.75pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: 1.0pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("		border-right: none;                                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1.0pt solid windowtext;                                                                                                                \n");
            htmlStr.Append("		border-left: 1.0pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl117309991 {                                                                                                                                           \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 1.0pt;                                                                                                                                    \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: 1.0pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("		border-right: none;                                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1.0pt solid windowtext;                                                                                                                \n");
            htmlStr.Append("		border-left: 1.0pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl11830999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 13.75pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: 1.0pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("		border-right: none;                                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1.0pt solid windowtext;                                                                                                                \n");
            htmlStr.Append("		border-left: none;                                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl11930999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: '_\\(* \\#\\,\\#\\#0_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                                    \n");
            htmlStr.Append("		text-align: right;                                                                                                                                   \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border: 1.0pt solid windowtext;                                                                                                                       \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl12030999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: '_\\(* \\#\\,\\#\\#0_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                                    \n");
            htmlStr.Append("		text-align: right;                                                                                                                                   \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: 1.0pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("		border-right: none;                                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1.0pt solid windowtext;                                                                                                                \n");
            htmlStr.Append("		border-left: 1.0pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl12130999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: 1.0pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("		border-right: none;                                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl12230999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: 1.0pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("		border-right: 1.0pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl12330999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1.0pt solid windowtext;                                                                                                                \n");
            htmlStr.Append("		border-left: 1.0pt solid windowtext;                                                                                                                  \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl12430999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1.0pt solid windowtext;                                                                                                                \n");
            htmlStr.Append("		border-left: none;                                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl12530999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                                    \n");
            htmlStr.Append("		border-right: 1.0pt solid windowtext;                                                                                                                 \n");
            htmlStr.Append("		border-bottom: 1.0pt solid windowtext;                                                                                                                \n");
            htmlStr.Append("		border-left: none;                                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl12630999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: red;                                                                                                                                          \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                                    \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1.0pt solid #0070C0;                                                                                                                   \n");
            htmlStr.Append("		border-left: 1.0pt solid #0070C0;                                                                                                                     \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl12730999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: red;                                                                                                                                          \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                                    \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1.0pt solid #0070C0;                                                                                                                   \n");
            htmlStr.Append("		border-left: none;                                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl12830999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: red;                                                                                                                                          \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                                    \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                                    \n");
            htmlStr.Append("		border-right: 1.0pt solid #0070C0;                                                                                                                    \n");
            htmlStr.Append("		border-bottom: 1.0pt solid #0070C0;                                                                                                                   \n");
            htmlStr.Append("		border-left: none;                                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl12930999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 14.0pt;;                                                                                                                                    \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1.0pt solid windowtext;                                                                                                                \n");
            htmlStr.Append("		border-left: none;                                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl13030999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 14.0pt;;                                                                                                                                    \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: 1.0pt solid windowtext;                                                                                                                   \n");
            htmlStr.Append("		border-right: none;                                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl13130999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                                        \n");
            htmlStr.Append("		font-size: 14.0pt;;                                                                                                                                    \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl13230999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: red;                                                                                                                                          \n");
            htmlStr.Append("		font-size: 13.75pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                                    \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: 1.0pt solid #0070C0;                                                                                                                      \n");
            htmlStr.Append("		border-right: none;                                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                                 \n");
            htmlStr.Append("		border-left: 1.0pt solid #0070C0;                                                                                                                     \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl13330999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: red;                                                                                                                                          \n");
            htmlStr.Append("		font-size: 13.75pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                                    \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: 1.0pt solid #0070C0;                                                                                                                      \n");
            htmlStr.Append("		border-right: none;                                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl13430999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: red;                                                                                                                                          \n");
            htmlStr.Append("		font-size: 13.75pt;                                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                                    \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: 1.0pt solid #0070C0;                                                                                                                      \n");
            htmlStr.Append("		border-right: 1.0pt solid #0070C0;                                                                                                                    \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl13530999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: red;                                                                                                                                          \n");
            htmlStr.Append("		font-size: 14.0pt;;                                                                                                                                    \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                                    \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                                 \n");
            htmlStr.Append("		border-left: 1.0pt solid #0070C0;                                                                                                                     \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl13630999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: red;                                                                                                                                          \n");
            htmlStr.Append("		font-size: 14.0pt;;                                                                                                                                    \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                                    \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                         \n");
            htmlStr.Append("	.xl13730999 {                                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                                 \n");
            htmlStr.Append("		color: red;                                                                                                                                          \n");
            htmlStr.Append("		font-size: 14.0pt;;                                                                                                                                    \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                                    \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                                    \n");
            htmlStr.Append("		border-right: 1.0pt solid #0070C0;                                                                                                                    \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                        \n");
            htmlStr.Append("	-->                                                                                                                                                      \n");
            htmlStr.Append("</style>                                                                                                                                                                                                \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append("</head>                                                                                                                                                                                                 \n");
            htmlStr.Append("<body class='A4'>                                                                                                                                                                                       \n");

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

            double v_totalHeightLastPage = 173.5;//203.5 243.5.5;

            double v_totalHeightPage = 570;//   540;

            for (int k = 0; k < v_countNumberOfPages; k++)
            {
                v_totalHeightPage = 510;// 330;

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

                htmlStr.Append("	<table border=0 cellpadding=0 cellspacing=0 width=683 class=xl6630999																																									 \n");
                htmlStr.Append("		style='border-collapse: collapse; table-layout: fixed; width: 510pt'>                                                                                                                                                                \n");
                htmlStr.Append("		<col class=xl6630999 width=7                                                                                                                                                                                                         \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 256; width: 6.35pt'>                                                                                                                                                               \n");
                htmlStr.Append("		<!-- 16 -->                                                                                                                                                                                                                          \n");
                htmlStr.Append("		<col class=xl6630999 width=6                                                                                                                                                                                                         \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 199; width: 5.0pt'>                                                                                                                                                               \n");
                htmlStr.Append("		<!-- 15 -->                                                                                                                                                                                                                          \n");
                htmlStr.Append("		<col class=xl6630999 width=26                                                                                                                                                                                                        \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 910; width: 23.75pt'>                                                                                                                                                              \n");
                htmlStr.Append("		<!-- 14 -->                                                                                                                                                                                                                          \n");
                htmlStr.Append("		<col class=xl6630999 width=55                                                                                                                                                                                                        \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 1962; width: 51.25pt'>                                                                                                                                                             \n");
                htmlStr.Append("		<!-- 13 -->                                                                                                                                                                                                                          \n");
                htmlStr.Append("		<col class=xl6630999 width=77                                                                                                                                                                                                        \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 2730; width: 97.5pt'>                                                                                                                                                             \n");
                htmlStr.Append("		<!-- 11 12 -->                                                                                                                                                                                                                       \n");
                htmlStr.Append("		<col class=xl6630999 width=62 span=2                                                                                                                                                                                                 \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 2218; width: 58.75pt'>                                                                                                                                                             \n");
                htmlStr.Append("		<!-- 8  9 10  -->                                                                                                                                                                                                                    \n");
                htmlStr.Append("		<col class=xl6630999 width=63 span=3                                                                                                                                                                                                 \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 2247; width: 65pt'>                                                                                                                                                             \n");
                htmlStr.Append("		<!-- 6 7  -->                                                                                                                                                                                                                        \n");
                htmlStr.Append("		<col class=xl6630999 width=42 span=2                                                                                                                                                                                                 \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 1479; width: 32.5pt'>                                                                                                                                                             \n");
                htmlStr.Append("		<!-- 5 -->                                                                                                                                                                                                                           \n");
                htmlStr.Append("		<col class=xl6630999 width=62                                                                                                                                                                                                        \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 2218; width: 58.75pt'>                                                                                                                                                             \n");
                htmlStr.Append("		<!-- 4 -->                                                                                                                                                                                                                           \n");
                htmlStr.Append("		<col class=xl6630999 width=38                                                                                                                                                                                                        \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 1336; width: 35.0pt'>                                                                                                                                                             \n");
                htmlStr.Append("		<!-- 3 -->                                                                                                                                                                                                                           \n");
                htmlStr.Append("		<col class=xl6630999 width=6                                                                                                                                                                                                         \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 199; width: 5.0pt'>                                                                                                                                                               \n");
                htmlStr.Append("		<!-- 2 -->                                                                                                                                                                                                                           \n");
                htmlStr.Append("		<col class=xl6630999 width=9                                                                                                                                                                                                         \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 312; width: 8.75pt'>                                                                                                                                                               \n");
                htmlStr.Append("		<!-- 1 -->                                                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr height=9 style='mso-height-source: userset; height: 9.7pt'>                                                                                                                                                                     \n");
                htmlStr.Append("			<td height=9 class=xl6330999 width=7                                                                                                                                                                                             \n");
                htmlStr.Append("				style='height: 9.7pt; width: 6.35pt'>&nbsp;</td>                                                                                                                                                                               \n");
                htmlStr.Append("			<td class=xl6430999 width=6 style='width: 5.0pt'>&nbsp;</td>                                                                                                                                                                       \n");
                htmlStr.Append("			<td class=xl6430999 width=26 style='width: 23.75pt'>&nbsp;</td>                                                                                                                                                                     \n");
                htmlStr.Append("			<td class=xl6430999 width=55 style='width: 51.25pt'>&nbsp;</td>                                                                                                                                                                     \n");
                htmlStr.Append("			<td class=xl6430999 width=77 style='width: 58pt'>&nbsp;</td>                                                                                                                                                                     \n");
                htmlStr.Append("			<td class=xl6430999 width=62 style='width: 58.75pt'>&nbsp;</td>                                                                                                                                                                     \n");
                htmlStr.Append("			<td class=xl6430999 width=62 style='width: 58.75pt'>&nbsp;</td>                                                                                                                                                                     \n");
                htmlStr.Append("			<td class=xl6430999 width=63 style='width: 58.75pt'>&nbsp;</td>                                                                                                                                                                     \n");
                htmlStr.Append("			<td class=xl6430999 width=63 style='width: 58.75pt'>&nbsp;</td>                                                                                                                                                                     \n");
                htmlStr.Append("			<td class=xl6430999 width=63 style='width: 58.75pt'>&nbsp;</td>                                                                                                                                                                     \n");
                htmlStr.Append("			<td class=xl6430999 width=42 style='width: 31pt'>&nbsp;</td>                                                                                                                                                                     \n");
                htmlStr.Append("			<td class=xl6430999 width=42 style='width: 31pt'>&nbsp;</td>                                                                                                                                                                     \n");
                htmlStr.Append("			<td class=xl6430999 width=62 style='width: 58.75pt'>&nbsp;</td>                                                                                                                                                                     \n");
                htmlStr.Append("			<td class=xl6430999 width=38 style='width: 35.0pt'>&nbsp;</td>                                                                                                                                                                     \n");
                htmlStr.Append("			<td class=xl6430999 width=6 style='width: 5.0pt'>&nbsp;</td>                                                                                                                                                                       \n");
                htmlStr.Append("			<td class=xl6530999 width=9 style='width: 8.75pt'>&nbsp;</td>                                                                                                                                                                       \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                                                                \n");
                htmlStr.Append("		<tr height=23 style='height: 17.4pt'>                                                                                                                                                                                                \n");
                htmlStr.Append("			<td height=23 class=xl6730999 style='height: 17.4pt'>&nbsp;</td>                                                                                                                                                                 \n");
                htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
                htmlStr.Append("			<td align=left valign=top><![if !vml]><span                                                                                                                                                                                      \n");
                htmlStr.Append("				style='mso-ignore: vglayout; position: absolute; z-index: 1; margin-left: -2px; margin-top: 5px; width: 80px; height: 73px'><img                                                                                             \n");
                htmlStr.Append("					width=80 height=73                                                                                                                                                                                                       \n");
                htmlStr.Append("					src='D:/webproject/e-invoice-ws/02.Web/EInvoice/img/OTM_001.png'                                                                                                                                                       \n");
            htmlStr.Append("					v:shapes='Picture_x0020_5'></span> <![endif]><span                                                                                                                                                                       \n");
            htmlStr.Append("				style='mso-ignore: vglayout2'>                                                                                                                                                                                               \n");
            htmlStr.Append("					<table cellpadding=0 cellspacing=0>                                                                                                                                                                                      \n");
            htmlStr.Append("						<tr>                                                                                                                                                                                                                 \n");
            htmlStr.Append("							<td height=23 class=xl6630999 width=26                                                                                                                                                                           \n");
            htmlStr.Append("								style='height: 17.4pt; width: 23.75pt'>&nbsp;</td>                                                                                                                                                              \n");
            htmlStr.Append("						</tr>                                                                                                                                                                                                                \n");
            htmlStr.Append("					</table>                                                                                                                                                                                                                 \n");
            htmlStr.Append("			</span></td>                                                                                                                                                                                                                     \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl10130999 colspan=5>" + dt.Rows[0]["Seller_name"] + "</td>                                                                                                                                                                            \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6830999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                                                                \n");
            htmlStr.Append("		<tr height=18 style='height: 17.25pt'>                                                                                                                                                                                                \n");
            htmlStr.Append("			<td height=18 class=xl6730999 style='height: 17.25pt'>&nbsp;</td>                                                                                                                                                                 \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6930999 colspan=9>&#272;&#7883;a ch&#7881; (<font                                                                                                                                                                    \n");
            htmlStr.Append("				class='font730999'>Address</font><font class='font630999'>):                                                                                                                                                                 \n");
            htmlStr.Append("					" + l_add + "</font></td>                                                                                                                                                                                                \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6830999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                                                                \n");
            htmlStr.Append("		<tr height=18 style='height: 17.25pt'>                                                                                                                                                                                                \n");
            htmlStr.Append("			<td height=18 class=xl6730999 style='height: 17.25pt'>&nbsp;</td>                                                                                                                                                                 \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6930999 colspan=4>" + l_add1 + "</td>                                                                                                                                                                                \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6830999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                                                                \n");
            htmlStr.Append("		<tr class=xl9430999 height=18 style='height: 17.25pt'>                                                                                                                                                                                \n");
            htmlStr.Append("			<td height=18 class=xl9330999 style='height: 17.25pt'>&nbsp;</td>                                                                                                                                                                 \n");
            htmlStr.Append("			<td class=xl9430999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl9430999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl9430999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl9530999 colspan=9>"+ l_adden +"</td>                                                                                                                                                                                      \n");
            htmlStr.Append("			<td class=xl9430999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl9430999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl9630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                                                                \n");
            htmlStr.Append("		<tr class=xl9430999 height=18 style='height: 17.25pt'>                                                                                                                                                                                \n");
            htmlStr.Append("			<td height=18 class=xl9330999 style='height: 17.25pt'>&nbsp;</td>                                                                                                                                                                 \n");
            htmlStr.Append("			<td class=xl9430999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl9430999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl9430999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl9430999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl9530999 colspan=7>"+ l_add1en+ "</td>                                                                                                                                                                                \n");
            htmlStr.Append("			<td class=xl9430999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl9430999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl9430999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl9630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                                                                \n");
            htmlStr.Append("		<tr height=18 style='height: 17.25pt'>                                                                                                                                                                                                \n");
            htmlStr.Append("			<td height=18 class=xl6730999 style='height: 17.25pt'>&nbsp;</td>                                                                                                                                                                 \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7030999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7030999 colspan=4>Mã s&#7889; thu&#7871; (<font                                                                                                                                                                      \n");
            htmlStr.Append("				class='font930999'>Tax code</font><font class='font830999'>):                                                                                                                                                                \n");
            htmlStr.Append("					" + dt.Rows[0]["Seller_TaxCode"] + "</font></td>                                                                                                                                                                                       \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6830999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                                                                \n");
            htmlStr.Append("		<tr height=18 style='height: 17.25pt'>                                                                                                                                                                                                \n");
            htmlStr.Append("			<td height=18 class=xl6730999 style='height: 17.25pt'>&nbsp;</td>                                                                                                                                                                 \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7030999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7030999 colspan=5>&#272;i&#7879;n tho&#7841;i (<font                                                                                                                                                                 \n");
            htmlStr.Append("				class='font930999'>Tel</font><font class='font830999'>): " + dt.Rows[0]["Seller_Tel"] + "                                                                                                                                                  \n");
            htmlStr.Append("					- Fax: " + dt.Rows[0]["Seller_Fax"] + "</font></td>                                                                                                                                                                                    \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6830999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                                                                \n");
            htmlStr.Append("		<tr height=9 style='mso-height-source: userset; height: 2.05pt'>                                                                                                                                                                     \n");
            htmlStr.Append("			<td height=9 class=xl6330999 style='height: 2.05pt'>&nbsp;</td>                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6430999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6430999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6430999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6430999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6430999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6430999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6430999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6430999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6430999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6430999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6430999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6430999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6430999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6430999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6530999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                                                                \n");
            htmlStr.Append("		<tr height=26 style='mso-height-source: userset; height: 25.1pt'>                                                                                                                                                                    \n");
            htmlStr.Append("			<td height=26 class=xl6730999 style='height: 25.1pt'>&nbsp;</td>                                                                                                                                                                 \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td colspan=8 class=xl10230999>PHI&#7870;U XU&#7844;T KHO KIÊM                                                                                                                                                                   \n");
            htmlStr.Append("				V&#7852;N CHUY&#7874;N N&#7896;I B&#7896;</td>                                                                                                                                                                               \n");
            htmlStr.Append("			<td class=xl6630999 colspan=5></td>                                                                                                                                                                                              \n");
            htmlStr.Append("			<td class=xl6830999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                                                                \n");
            htmlStr.Append("		<tr height=26 style='mso-height-source: userset; height: 25.1pt'>                                                                                                                                                                    \n");
            htmlStr.Append("			<td height=26 class=xl6730999 style='height: 25.1pt'>&nbsp;</td>                                                                                                                                                                 \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl10230999>&nbsp;</td>                                                                                                                                                                                                 \n");
            htmlStr.Append("			<td class=xl10230999>&nbsp;</td>                                                                                                                                                                                                 \n");
            htmlStr.Append("			<td class=xl10030999 colspan=3>DELIVERY NOTE CUM INTERNAL                                                                                                                                                                        \n");
            htmlStr.Append("				TRANSPORT</td>                                                                                                                                                                                                               \n");
            htmlStr.Append("			<td class=xl10230999>&nbsp;</td>                                                                                                                                                                                                 \n");
            htmlStr.Append("			<td class=xl10230999>&nbsp;</td>                                                                                                                                                                                                 \n");
            htmlStr.Append("			<td class=xl10230999>&nbsp;</td>                                                                                                                                                                                                 \n");
            htmlStr.Append("			<td class=xl6630999 colspan=4><font class='font830999'>Ký                                                                                                                                                                        \n");
            htmlStr.Append("					hi&#7879;u (</font><font class='font930999'>Serial</font><font                                                                                                                                                           \n");
            htmlStr.Append("				class='font830999'>):</font><font class='font530999'> </font><font                                                                                                                                                           \n");
            htmlStr.Append("				class='font1030999'>" + dt.Rows[0]["templateCode"] + "" + dt.Rows[0]["InvoiceSerialNo"] + "</font></td>                                                                                                                                                               \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6830999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                                                                \n");
            htmlStr.Append("		<tr height=22 style='height: 21.0pt'>                                                                                                                                                                                                \n");
            htmlStr.Append("			<td height=22 class=xl6730999 style='height: 21.0pt'>&nbsp;</td>                                                                                                                                                                 \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td colspan=8 class=xl10730999>Ngày (<font class='font1530999'>date</font><font                                                                                                                                                  \n");
            htmlStr.Append("				class='font530999'>)<span style='mso-spacerun: yes'> " + dt.Rows[0]["invoiceissueddate_dd"] + "                                                                                                                                                          \n");
            htmlStr.Append("				</span>tháng (                                                                                                                                                                                                               \n");
            htmlStr.Append("			</font><font class='font1530999'>month</font><font class='font530999'>)<span                                                                                                                                                     \n");
            htmlStr.Append("					style='mso-spacerun: yes'> " + dt.Rows[0]["invoiceissueddate_mm"] + "                                                                                                                                                                                \n");
            htmlStr.Append("				</span>n&#259;m (                                                                                                                                                                                                            \n");
            htmlStr.Append("			</font><font class='font1530999'>year</font><font class='font530999'>)<span                                                                                                                                                      \n");
            htmlStr.Append("					style='mso-spacerun: yes'> " + dt.Rows[0]["invoiceissueddate_yyyy"] + "                                                                                                                                                                              \n");
            htmlStr.Append("				</span></font></td>                                                                                                                                                                                                          \n");
            htmlStr.Append("			<td class=xl6630999 colspan=3><font class='font830999'>S&#7889;                                                                                                                                                                  \n");
            htmlStr.Append("					(</font><font class='font930999'>No</font><font class='font830999'>.):</font><font                                                                                                                                       \n");
            htmlStr.Append("				class='font530999'> </font><font class='font1130999'>" + dt.Rows[0]["InvoiceNumber"] + "</font></td>                                                                                                                                          \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6830999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                                                                \n");
            htmlStr.Append("	                                                                                                                                                                                                                                         \n");
            htmlStr.Append("		<tr height=18 style='height: 2.8pt'>                                                                                                                                                                                                 \n");
            htmlStr.Append("			<td height=18 class=xl67309991                                                                                                                                                                                                   \n");
            htmlStr.Append("				style='height: 2.8pt; border-left: 1.0pt solid windowtext;'>&nbsp;</td>                                                                                                                                                       \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6830999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                                                                \n");
            htmlStr.Append("	                                                                                                                                                                                                                                         \n");
            htmlStr.Append("		<tr class=xl7330999 height=22                                                                                                                                                                                                        \n");
            htmlStr.Append("			style='mso-height-source: userset; height: 21.38pt'>                                                                                                                                                                              \n");
            htmlStr.Append("			<td height=22 class=xl7130999                                                                                                                                                                                                    \n");
            htmlStr.Append("				style='height: 21.38pt; border-right: 1.0pt solid windowtext;'>&nbsp;</td>                                                                                                                                                     \n");
            htmlStr.Append("			<td class=xl7130999                                                                                                                                                                                                              \n");
            htmlStr.Append("				style='height: 21.38pt; border-top: 1.0pt solid windowtext;'>&nbsp;</td>                                                                                                                                                       \n");
            htmlStr.Append("			<td class=xl7030999 colspan=5                                                                                                                                                                                                    \n");
            htmlStr.Append("				style='height: 21.38pt; border-top: 1.0pt solid windowtext;'>C&#259;n                                                                                                                                                          \n");
            htmlStr.Append("				c&#7913; l&#7879;nh &#273;i&#7873;u &#273;&#7897;ng s&#7889; (<font                                                                                                                                                          \n");
            htmlStr.Append("				class='font930999'>Ordering no</font><font class='font830999'>):                                                                                                                                                             \n");
            htmlStr.Append("					" + dt.Rows[0]["attribute_01"] + "</font>                                                                                                                                                                                               \n");
            htmlStr.Append("			</td>                                                                                                                                                                                                                            \n");
            htmlStr.Append("			<td class=xl7030999                                                                                                                                                                                                              \n");
            htmlStr.Append("				style='height: 21.38pt; border-top: 1.0pt solid windowtext;'>&nbsp;</td>                                                                                                                                                       \n");
            htmlStr.Append("			<td class=xl7030999 colspan=6                                                                                                                                                                                                    \n");
            htmlStr.Append("				style='height: 21.38pt; border-top: 1.0pt solid windowtext;'>Ngày                                                                                                                                                              \n");
            htmlStr.Append("				(<font class='font930999'>date</font><font class='font830999'>):                                                                                                                                                             \n");
            htmlStr.Append("			</font>" + dt.Rows[0]["attribute_02"] + "</td>                                                                                                                                                                                                  \n");
            //htmlStr.Append("			<td class=xl7030999                                                                                                                                                                                                              \n");
            //htmlStr.Append("				style='height: 21.38pt; border-top: 1.0pt solid windowtext;'>&nbsp;</td>                                                                                                                                                       \n");
            //htmlStr.Append("			<td class=xl7030999                                                                                                                                                                                                              \n");
            //htmlStr.Append("				style='height: 21.38pt; border-top: 1.0pt solid windowtext;'>&nbsp;</td>                                                                                                                                                       \n");
            //htmlStr.Append("			<td class=xl7030999                                                                                                                                                                                                              \n");
            //htmlStr.Append("				style='height: 21.38pt; border-top: 1.0pt solid windowtext;'>&nbsp;</td>                                                                                                                                                       \n");
            //htmlStr.Append("			<td class=xl7030999                                                                                                                                                                                                              \n");
            //htmlStr.Append("				style='height: 21.38pt; border-top: 1.0pt solid windowtext;'>&nbsp;</td>                                                                                                                                                       \n");
            htmlStr.Append("			<td class=xl7230999                                                                                                                                                                                                              \n");
            htmlStr.Append("				style='height: 21.38pt; border-top: 1.0pt solid windowtext;'>&nbsp;</td>                                                                                                                                                       \n");
            htmlStr.Append("			<td class=xl7230999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                                                                \n");
            htmlStr.Append("	                                                                                                                                                                                                                                         \n");
            htmlStr.Append("		<tr class=xl7330999 height=22                                                                                                                                                                                                        \n");
            htmlStr.Append("			style='mso-height-source: userset; height: 21.38pt'>                                                                                                                                                                              \n");
            htmlStr.Append("			<td height=22 class=xl7130999 style='height: 21.38pt'>&nbsp;</td>                                                                                                                                                                 \n");
            htmlStr.Append("			<td class=xl7130999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7030999 colspan=6>C&#7911;a (<font                                                                                                                                                                                   \n");
            htmlStr.Append("				class='font930999'>By</font><font class='font830999'>): </font>" + dt.Rows[0]["attribute_03"] + "</td>                                                                                                                                      \n");
            htmlStr.Append("			<td class=xl7030999 colspan=7 style='border-right:01.0pt solid windowtext'>v&#7873; vi&#7879;c (<font                                                                                                                             \n");
            htmlStr.Append("				class='font930999'>for</font><font class='font830999'>): </font> " + dt.Rows[0]["attribute_04"] + "</td>                                                                                                                                    \n");
            htmlStr.Append("			<td class=xl6830999 style='border-left:01.0pt solid windowtext'>&nbsp;</td>                                                                                                                                                       \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                                                                \n");
            htmlStr.Append("		<tr height=22 style='mso-height-source: userset; height: 21.38pt'>                                                                                                                                                                    \n");
            htmlStr.Append("			<td height=22 class=xl6730999 style='height: 21.38pt'>&nbsp;</td>                                                                                                                                                                 \n");
            htmlStr.Append("			<td class=xl6730999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7030999 colspan=6>H&#7885; tên ng&#432;&#7901;i                                                                                                                                                                      \n");
            htmlStr.Append("				v&#7853;n chuy&#7875;n (<font class='font930999'>Transporter</font><font class='font830999'>):</font> " + dt.Rows[0]["attribute_05"] + "                                                                                                    \n");
            htmlStr.Append("			</td>                                                                                                                                                                                                                            \n");
            htmlStr.Append("			<td class=xl7030999 colspan=7 style='border-right:01.0pt solid windowtext'>H&#7907;p &#273;&#7891;ng s&#7889;                                                                                                                     \n");
            htmlStr.Append("				(<font class='font930999'>Contract no</font><font class='font830999'>):                                                                                                                                                      \n");
            htmlStr.Append("			</font> " + dt.Rows[0]["attribute_06"] + "                                                                                                                                                                                                      \n");
            htmlStr.Append("			</td>                                                                                                                                                                                                                            \n");
            htmlStr.Append("			                                                                                                                                                                                                                                 \n");
            htmlStr.Append("			<td class=xl6830999 style='border-left:01.0pt solid windowtext'>&nbsp;</td>                                                                                                                                                       \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                                                                \n");
            htmlStr.Append("		<tr height=22 style='mso-height-source: userset; height: 21.38pt'>                                                                                                                                                                    \n");
            htmlStr.Append("			<td height=22 class=xl6730999 style='height: 21.38pt'>&nbsp;</td>                                                                                                                                                                 \n");
            htmlStr.Append("			<td class=xl6730999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7030999 colspan=13 style='border-right:01.0pt solid windowtext'>Ph&#432;&#417;ng ti&#7879;n                                                                                                                           \n");
            htmlStr.Append("				v&#7853;n t&#7843;i (<font class='font930999'>Transportation</font><font                                                                                                                                                     \n");
            htmlStr.Append("				class='font830999'>): </font> " + dt.Rows[0]["attribute_07"] + "                                                                                                                                                                            \n");
            htmlStr.Append("			</td>                                                                                                                                                                                                                            \n");
            htmlStr.Append("			                                                                                                                                                                                                                                 \n");
            htmlStr.Append("			<td class=xl6830999 style='border-left:01.0pt solid windowtext'>&nbsp;</td>                                                                                                                                                       \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                                                                \n");
            htmlStr.Append("		<tr height=22 style='mso-height-source: userset; height: 21.38pt'>                                                                                                                                                                    \n");
            htmlStr.Append("			<td height=22 class=xl6730999 style='height: 21.38pt'>&nbsp;</td>                                                                                                                                                                 \n");
            htmlStr.Append("			<td class=xl6730999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7030999 colspan=3>Xu&#7845;t t&#7841;i kho (<font                                                                                                                                                                    \n");
            htmlStr.Append("				class='font930999'>Stock out at</font><font class='font830999'>):</font></td>                                                                                                                                                \n");
            htmlStr.Append("			<td class=xl7430999 colspan=4>" + dt.Rows[0]["attribute_08"] + "</td>                                                                                                                                                                           \n");
            htmlStr.Append("			<td class=xl7030999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7530999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7530999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7030999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7030999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6830999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6830999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                                                                \n");
            htmlStr.Append("		<tr height=22 style='mso-height-source: userset; height: 21.38pt'>                                                                                                                                                                    \n");
            htmlStr.Append("			<td height=22 class=xl6730999 style='height: 21.38pt'>&nbsp;</td>                                                                                                                                                                 \n");
            htmlStr.Append("			<td class=xl6730999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7030999 colspan=3>Nh&#7853;p t&#7841;i kho (<font                                                                                                                                                                    \n");
            htmlStr.Append("				class='font930999'>Stock in at</font><font class='font830999'>):</font></td>                                                                                                                                                 \n");
            htmlStr.Append("			<td class=xl7430999 colspan=4>" + dt.Rows[0]["attribute_09"] + "</td>                                                                                                                                                                           \n");
            htmlStr.Append("			<td class=xl7030999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7530999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7530999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7030999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7030999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6830999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6830999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                                                                \n");
            htmlStr.Append("	                                                                                                                                                                                                                                         \n");
            htmlStr.Append("		<tr class=xl7330999 height=29                                                                                                                                                                                                        \n");
            htmlStr.Append("			style='mso-height-source: userset; height: 27.6pt'>                                                                                                                                                                             \n");
            htmlStr.Append("			<td height=29 class=xl7130999 style='height: 27.6pt'>&nbsp;</td>                                                                                                                                                                \n");
            htmlStr.Append("			<td colspan=2 rowspan=2 class=xl10630999 width=32 style='width: 23pt'>STT<br>                                                                                                                                                    \n");
            htmlStr.Append("				<font class='font1630999'>No</font></td>                                                                                                                                                                                     \n");
            htmlStr.Append("			<td colspan=3 rowspan=2 class=xl10930999 width=194                                                                                                                                                                               \n");
            htmlStr.Append("				style='border-right: 1.0pt solid black; border-bottom: 1.0pt solid black; width: 146pt'>Tên                                                                                                                                    \n");
            htmlStr.Append("				hàng hóa, quy cách<br> ph&#7849;m ch&#7845;t v&#7853;t t&#432;                                                                                                                                                               \n");
            htmlStr.Append("				(S&#7843;n ph&#7849;m hàng hóa)<br> (<font class='font1630999'>Product                                                                                                                                                       \n");
            htmlStr.Append("					name, specification)</font>                                                                                                                                                                                              \n");
            htmlStr.Append("			</td>                                                                                                                                                                                                                            \n");
            htmlStr.Append("			<td rowspan=2 class=xl10630999 width=62 style='width: 58.75pt'>Mã                                                                                                                                                                   \n");
            htmlStr.Append("				s&#7889;<br> (<font class='font1630999'>Part no)</font>                                                                                                                                                                      \n");
            htmlStr.Append("			</td>                                                                                                                                                                                                                            \n");
            htmlStr.Append("			<td rowspan=2 class=xl10630999 width=63 style='width: 58.75pt'>&#272;&#417;n                                                                                                                                                        \n");
            htmlStr.Append("				v&#7883; <br> tính<br> (<font class='font1630999'>Unit)</font>                                                                                                                                                               \n");
            htmlStr.Append("			</td>                                                                                                                                                                                                                            \n");
            htmlStr.Append("			<td colspan=2 class=xl10830999 style='border-left: none'>S&#7889;                                                                                                                                                                \n");
            htmlStr.Append("				l&#432;&#7907;ng (<font class='font1630999'>Q'ty</font><font                                                                                                                                                                 \n");
            htmlStr.Append("				class='font1230999'>)</font>                                                                                                                                                                                                 \n");
            htmlStr.Append("			</td>                                                                                                                                                                                                                            \n");
            htmlStr.Append("			<td colspan=2 rowspan=2 class=xl10930999 width=84                                                                                                                                                                                \n");
            htmlStr.Append("				style='border-bottom: 1.0pt solid black; width: 62pt'>&#272;&#417;n                                                                                                                                                           \n");
            htmlStr.Append("				giá<br> (<font class='font1630999'>Unit price</font><font                                                                                                                                                                    \n");
            htmlStr.Append("				class='font1230999'>)</font>                                                                                                                                                                                                 \n");
            htmlStr.Append("			</td>                                                                                                                                                                                                                            \n");
            htmlStr.Append("			<td colspan=3 rowspan=2 class=xl10630999 width=106                                                                                                                                                                               \n");
            htmlStr.Append("				style='width: 79pt'>Thành ti&#7873;n<br> (<font                                                                                                                                                                              \n");
            htmlStr.Append("				class='font1630999'>Amount</font><font class='font1230999'>)</font></td>                                                                                                                                                     \n");
            htmlStr.Append("			<td class=xl7230999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                                                                \n");
            htmlStr.Append("		<tr class=xl7930999 height=62                                                                                                                                                                                                        \n");
            htmlStr.Append("			style='mso-height-source: userset; height: 58.2pt'>                                                                                                                                                                              \n");
            htmlStr.Append("			<td height=62 class=xl7730999 style='height: 58.2pt'>&nbsp;</td>                                                                                                                                                                 \n");
            htmlStr.Append("			<td class=xl10630999 width=63                                                                                                                                                                                                    \n");
            htmlStr.Append("				style='border-top: none; border-left: none; width: 58.75pt'>Th&#7921;c<br>                                                                                                                                                      \n");
            htmlStr.Append("				xu&#7845;t<br> (<font class='font1630999'>Out)</font></td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl10630999 width=63                                                                                                                                                                                                    \n");
            htmlStr.Append("				style='border-top: none; border-left: none; width: 58.75pt'>Th&#7921;c<br>                                                                                                                                                      \n");
            htmlStr.Append("				nh&#7853;p<br> (<font class='font1630999'>In)</font></td>                                                                                                                                                                    \n");
            htmlStr.Append("			<td class=xl7830999>&nbsp;</td>                                                                                                                                                                                                  \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                                                                \n");

            v_rowHeight = "26.0pt"; //"26.5pt";
                v_rowHeightEmpty = "22.0pt";
                v_rowHeightNumber = 26;

                v_rowHeightLast = "21.0pt";// "23.5pt";
                v_rowHeightLastNumber = 26;// 23.5;
                v_rowHeightEmptyLast = "23.5pt"; //"23.5pt";


                for (int dtR = 0; dtR < page[k]; dtR++)
                {
                    if (!vlongItemName && dt_d.Rows[v_index]["ITEM_NAME"].ToString().Length >= 92)
                    {
                        v_rowHeight = "26.0pt"; //"26.5pt";    
                        v_rowHeightLast = "21.0pt"; //"27.5pt";
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
               
                        htmlStr.Append("	            <tr height=24 style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>																	  \n");
                        htmlStr.Append("	            	<td height=24 class=xl6730999 style='height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>&nbsp;</td>                                                                 \n");
                        htmlStr.Append("	            	<td colspan=2 class=xl10330999>&nbsp;" + dt_d.Rows[dtR][7] + "</td>                                                                           \n");
                        htmlStr.Append("	            	<td colspan=3 class=xl11330999                                                                                                    \n");
                        htmlStr.Append("	            		style='border-right: 1.0pt solid black; border-left: none'>&nbsp;" + dt_d.Rows[dtR][0] + "</td>                                            \n");
                        htmlStr.Append("	            	<td class=xl8030999 style='border-top: none; border-left: none'>&nbsp;" + dt_d.Rows[dtR][8] + "</td>                                          \n");
                        htmlStr.Append("	            	<td class=xl10330999 style='border-top: none; border-left: none'>" + dt_d.Rows[dtR][1] + "&nbsp;</td>                                         \n");
                        htmlStr.Append("	            	<td class=xl8130999 style='border-top: none; border-left: none'>" + dt_d.Rows[dtR][2] + "&nbsp;</td>                                          \n");
                        htmlStr.Append("	            	<td class=xl8230999 style='border-top: none'>&nbsp;</td>                                                                          \n");
                        htmlStr.Append("	            	<td colspan=2 class=xl10430999 style='border-left: none;border-top: 1.0pt solid windowtext'>" + dt_d.Rows[dtR][3] + "&nbsp;</td>                                                 \n");
                        htmlStr.Append("	            	<td colspan=2 class=xl10430999 style='border-top: 1.0pt solid windowtext'>" + dt_d.Rows[dtR][4] + "&nbsp;</td>                                                                           \n");
                        htmlStr.Append("	            	<td class=xl8330999 style='border-top: none'>&nbsp;</td>                                                                          \n");
                        htmlStr.Append("	            	<td class=xl6830999>&nbsp;</td>                                                                                                   \n");
                        htmlStr.Append("	            </tr>                                                                                                                                 \n");
                    }
                    else if (dtR == page[k] - 1)//dong cuoi moi trang
                    {
                        if (k < v_countNumberOfPages - 1) //trang giua
                        {
                            
                            htmlStr.Append("		<tr height=24 style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>													   \n");
                            htmlStr.Append("			<td height=24 class=xl6730999 style='height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>&nbsp;</td>                                                   \n");
                            htmlStr.Append("			<td colspan=2 class=xl10330999>&nbsp;" + dt_d.Rows[dtR][7] + "</td>                                                            \n");
                            htmlStr.Append("			<td colspan=3 class=xl11330999                                                                                     \n");
                            htmlStr.Append("				style='border-right: 1.0pt solid black; border-left: none'>&nbsp;" + dt_d.Rows[dtR][0] + "</td>                             \n");
                            htmlStr.Append("			<td class=xl8030999 style='border-top: none; border-left: none'>&nbsp;" + dt_d.Rows[dtR][8] + "</td>                           \n");
                            htmlStr.Append("			<td class=xl10330999 style='border-top: none; border-left: none'>&nbsp;" + dt_d.Rows[dtR][1] + "</td>                          \n");
                            htmlStr.Append("			<td class=xl8130999 style='border-top: none; border-left: none'>" + dt_d.Rows[dtR][2] + "&nbsp;</td>                           \n");
                            htmlStr.Append("			<td class=xl8230999 style='border-top: none'>&nbsp;</td>                                                           \n");
                            htmlStr.Append("			<td colspan=2 class=xl10430999 style='border-left: none'>" + dt_d.Rows[dtR][3] + "&nbsp;</td>                                  \n");
                            htmlStr.Append("			<td colspan=2 class=xl10430999>" + dt_d.Rows[dtR][4] + "&nbsp;</td>                                                            \n");
                            htmlStr.Append("			<td class=xl8330999 style='border-top: none'>&nbsp;</td>                                                           \n");
                            htmlStr.Append("			<td class=xl6830999>&nbsp;</td>                                                                                    \n");
                            htmlStr.Append("		</tr>                                                                                                                  \n");


                        }
                        else // trang cuoi
                        {
                            if (dtR == rowsPerPage - 1) // du 11 dong
                            {
                    
                                htmlStr.Append("		<tr height=24 style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>													   \n");
                                htmlStr.Append("			<td height=24 class=xl6730999 style='height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>&nbsp;</td>                                                   \n");
                                htmlStr.Append("			<td colspan=2 class=xl10330999>&nbsp;" + dt_d.Rows[dtR][7] + "</td>                                                            \n");
                                htmlStr.Append("			<td colspan=3 class=xl11330999                                                                                     \n");
                                htmlStr.Append("				style='border-right: 1.0pt solid black; border-left: none'>&nbsp;" + dt_d.Rows[dtR][0] + "</td>                             \n");
                                htmlStr.Append("			<td class=xl8030999 style='border-top: none; border-left: none'>&nbsp;" + dt_d.Rows[dtR][8] + "</td>                           \n");
                                htmlStr.Append("			<td class=xl10330999 style='border-top: none; border-left: none'>&nbsp;" + dt_d.Rows[dtR][1] + "</td>                          \n");
                                htmlStr.Append("			<td class=xl8130999 style='border-top: none; border-left: none'>" + dt_d.Rows[dtR][2] + "&nbsp;</td>                           \n");
                                htmlStr.Append("			<td class=xl8230999 style='border-top: none'>&nbsp;</td>                                                           \n");
                                htmlStr.Append("			<td colspan=2 class=xl10430999 style='border-left: none'>" + dt_d.Rows[dtR][3] + "&nbsp;</td>                                  \n");
                                htmlStr.Append("			<td colspan=2 class=xl10430999>" + dt_d.Rows[dtR][4] + "&nbsp;</td>                                                            \n");
                                htmlStr.Append("			<td class=xl8330999 style='border-top: none'>&nbsp;</td>                                                           \n");
                                htmlStr.Append("			<td class=xl6830999>&nbsp;</td>                                                                                    \n");
                                htmlStr.Append("		</tr>                                                                                                                  \n");

                            }
                            else
                            {
                              
                                htmlStr.Append("		<tr height=24 style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>													   \n");
                                htmlStr.Append("			<td height=24 class=xl6730999 style='height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>&nbsp;</td>                                                   \n");
                                htmlStr.Append("			<td colspan=2 class=xl10330999>&nbsp;" + dt_d.Rows[dtR][7] + "</td>                                                            \n");
                                htmlStr.Append("			<td colspan=3 class=xl11330999                                                                                     \n");
                                htmlStr.Append("				style='border-right: 1.0pt solid black; border-left: none'>&nbsp;" + dt_d.Rows[dtR][0] + "</td>                             \n");
                                htmlStr.Append("			<td class=xl8030999 style='border-top: none; border-left: none'>&nbsp;" + dt_d.Rows[dtR][8] + "</td>                           \n");
                                htmlStr.Append("			<td class=xl10330999 style='border-top: none; border-left: none'>&nbsp;" + dt_d.Rows[dtR][1] + "</td>                          \n");
                                htmlStr.Append("			<td class=xl8130999 style='border-top: none; border-left: none'>" + dt_d.Rows[dtR][2] + "&nbsp;</td>                           \n");
                                htmlStr.Append("			<td class=xl8230999 style='border-top: none'>&nbsp;</td>                                                           \n");
                                htmlStr.Append("			<td colspan=2 class=xl10430999 style='border-left: none'>" + dt_d.Rows[dtR][3] + "&nbsp;</td>                                  \n");
                                htmlStr.Append("			<td colspan=2 class=xl10430999>" + dt_d.Rows[dtR][4] + "&nbsp;</td>                                                            \n");
                                htmlStr.Append("			<td class=xl8330999 style='border-top: none'>&nbsp;</td>                                                           \n");
                                htmlStr.Append("			<td class=xl6830999>&nbsp;</td>                                                                                    \n");
                                htmlStr.Append("		</tr>                                                                                                                  \n");


                            }

                        }
                    }
                    else
                    { // dong giua                                                                                                                                    
                       
                        htmlStr.Append("	<tr height=24 style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>																\n");
                        htmlStr.Append("		<td height=24 class=xl6730999 style='height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>&nbsp;</td>                                                            \n");
                        htmlStr.Append("		<td colspan=2 class=xl10330999>&nbsp;" + dt_d.Rows[dtR][7] + "</td>                                                                     \n");
                        htmlStr.Append("		<td colspan=3 class=xl11330999                                                                                              \n");
                        htmlStr.Append("			style='border-right: 1.0pt solid black; border-left: none'>&nbsp;" + dt_d.Rows[dtR][0] + "</td>                                      \n");
                        htmlStr.Append("		<td class=xl8030999 style='border-top: none; border-left: none'>&nbsp;" + dt_d.Rows[dtR][8] + "</td>                                    \n");
                        htmlStr.Append("		<td class=xl10330999 style='border-top: none; border-left: none'>" + dt_d.Rows[dtR][1] + "</td>                                         \n");
                        htmlStr.Append("		<td class=xl8130999 style='border-top: none; border-left: none'>" + dt_d.Rows[dtR][2] + "&nbsp;</td>                                    \n");
                        htmlStr.Append("		<td class=xl8230999 style='border-top: none'>&nbsp;</td>                                                                    \n");
                        htmlStr.Append("		<td colspan=2 class=xl10430999 style='border-left: none'>" + dt_d.Rows[dtR][3] + "&nbsp;</td>                                           \n");
                        htmlStr.Append("		<td colspan=2 class=xl10430999>" + dt_d.Rows[dtR][4] + "&nbsp;</td>                                                                     \n");
                        htmlStr.Append("		<td class=xl8330999 style='border-top: none'>&nbsp;</td>                                                                    \n");
                        htmlStr.Append("		<td class=xl6830999>&nbsp;</td>                                                                                             \n");
                        htmlStr.Append("	</tr>                                                                                                                           \n");
                    
                    }
                    v_index++;
                } //for dtR

                v_spacePerPage = 0;
                if (k < v_countNumberOfPages - 1 && page_index[k] < rowsPerPage)
                {
                    //for (int i = 0; i < rowsPerPage - page[k]; i++)
                    //{
                    v_spacePerPage += v_totalHeightPage;
                    //}
                }
                else if (k < v_countNumberOfPages - 1 && page_index[k] == rowsPerPage)
                {
                    v_spacePerPage = 18;
                }

                if (k == v_countNumberOfPages - 1 && page_index[k] < rowsPerPage) // Trang cuoi khong du dong
                {
                    v_rowHeightEmptyLast = Math.Round(v_totalHeightLastPage / (rowsPerPage - page_index[k]), 2).ToString() + "pt";
                    for (int i = 0; i < rowsPerPage - page_index[k]; i++)
                    {
                        if (i == (rowsPerPage - page_index[k] - 1))
                        {

                            htmlStr.Append("		<tr height=24 style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>													   \n");
                            htmlStr.Append("			<td height=24 class=xl6730999 style='height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>&nbsp;</td>                                                   \n");
                            htmlStr.Append("			<td colspan=2 class=xl10330999>&nbsp;</td>                                                            \n");
                            htmlStr.Append("			<td colspan=3 class=xl11330999                                                                                     \n");
                            htmlStr.Append("				style='border-right: 1.0pt solid black; border-left: none'>&nbsp;</td>                             \n");
                            htmlStr.Append("			<td class=xl8030999 style='border-top: none; border-left: none'>&nbsp;</td>                           \n");
                            htmlStr.Append("			<td class=xl10330999 style='border-top: none; border-left: none'>&nbsp;</td>                          \n");
                            htmlStr.Append("			<td class=xl8130999 style='border-top: none; border-left: none'>&nbsp;</td>                           \n");
                            htmlStr.Append("			<td class=xl8230999 style='border-top: none'>&nbsp;</td>                                                           \n");
                            htmlStr.Append("			<td colspan=2 class=xl10430999 style='border-left: none'>&nbsp;</td>                                  \n");
                            htmlStr.Append("			<td colspan=2 class=xl10430999>&nbsp;</td>                                                            \n");
                            htmlStr.Append("			<td class=xl8330999 style='border-top: none'>&nbsp;</td>                                                           \n");
                            htmlStr.Append("			<td class=xl6830999>&nbsp;</td>                                                                                    \n");
                            htmlStr.Append("		</tr>                                                                                                                  \n");

                        }
                        else
                        {
                            htmlStr.Append("	<tr height=24 style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>																\n");
                            htmlStr.Append("		<td height=24 class=xl6730999 style='height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>&nbsp;</td>                                                            \n");
                            htmlStr.Append("		<td colspan=2 class=xl10330999>&nbsp;</td>                                                                     \n");
                            htmlStr.Append("		<td colspan=3 class=xl11330999                                                                                              \n");
                            htmlStr.Append("			style='border-right: 1.0pt solid black; border-left: none'>&nbsp;</td>                                      \n");
                            htmlStr.Append("		<td class=xl8030999 style='border-top: none; border-left: none'>&nbsp;</td>                                    \n");
                            htmlStr.Append("		<td class=xl10330999 style='border-top: none; border-left: none'></td>                                         \n");
                            htmlStr.Append("		<td class=xl8130999 style='border-top: none; border-left: none'>&nbsp;</td>                                    \n");
                            htmlStr.Append("		<td class=xl8230999 style='border-top: none'>&nbsp;</td>                                                                    \n");
                            htmlStr.Append("		<td colspan=2 class=xl10430999 style='border-left: none'>&nbsp;</td>                                           \n");
                            htmlStr.Append("		<td colspan=2 class=xl10430999>&nbsp;</td>                                                                     \n");
                            htmlStr.Append("		<td class=xl8330999 style='border-top: none'>&nbsp;</td>                                                                    \n");
                            htmlStr.Append("		<td class=xl6830999>&nbsp;</td>                                                                                             \n");
                            htmlStr.Append("	</tr>                                                                                                                           \n");

                        }
                    } // for

                }//Trang cuoi 11 dong

                if (k < v_countNumberOfPages - 1)
                {

                    htmlStr.Append("       <tr class=xl767652 height = 25   \n");
                    htmlStr.Append("                    style='mso-height-source: userset; height: 18pt;border-bottom: 1.0pt solid black;' >                                                               \n");
                    htmlStr.Append("    				<td colspan = 2 height=25 class=xl767653 width = 39  \n");
                    htmlStr.Append("                        style=' height:18pt; width: 29pt'>&nbsp;</td>                                                               \n");
                    htmlStr.Append("    				<td colspan = 5 class=xl767653 width = 166 style='width: 125pt;border-right: none'>&nbsp;</td>                                                               \n");
                    htmlStr.Append("    				<td class=xl767653 width = 78 style='width: 55.1pt;border-left: none;border-right: none'>&nbsp;</td>                                                               \n");
                    htmlStr.Append("    				<td colspan = 2 class=xl767653                                                               \n");
                    htmlStr.Append("                          style = ' width: 21.85pt' > &nbsp;</td>                                                               \n");
                    htmlStr.Append("    				<td class=xl767653 width = 6 style=' width: 3.8pt;border-right: none'>&nbsp;</td>                                                               \n");
                    htmlStr.Append("    				<td colspan = 2 class=xl767653 style = 'border-left: none;border-right: none' > &nbsp;</td>                                                               \n");
                    htmlStr.Append("    				<td class=xl767653 width = 6 style=' width: 3.8pt;border-right: none'>&nbsp;</td>                                                               \n");
                    htmlStr.Append("    				<td colspan = 2 class=xl767653 style = 'border-left: none;border-right: none' > &nbsp;</td>                                                               \n");
                    htmlStr.Append("    				<td class=xl767653 width = 6 ></ td >                                                               \n");
                    htmlStr.Append("                                                                   \n");
                    htmlStr.Append("                </tr>                                                               \n");

                   

                    htmlStr.Append("	<table  border=0>                                                                                                                                                                                                 \n");
                    htmlStr.Append("		<tr height=5  style='height: " + (v_spacePerPage).ToString() + "pt'>                                                                                                                                                                \n");

                    htmlStr.Append("		</tr>      																																														\n");
                    htmlStr.Append("	</table>             																																										\n");

                }


            }// for k                                                                                                                             

            htmlStr.Append("		<tr height=29 style='mso-height-source: userset; height: 27.6pt'>																					 \n");
            htmlStr.Append("			<td height=29 class=xl6730999 style='height: 27.6pt'>&nbsp;</td>                                                                                \n");
            htmlStr.Append("			<td colspan=11 class=xl11730999>T&#7893;ng c&#7897;ng (<font                                                                                     \n");
            htmlStr.Append("				class='font1530999'>Total</font><font class='font530999'>)</font></td>                                                                       \n");
            htmlStr.Append("			<td colspan=3 class=xl11930999>" + dt.Rows[0]["TOT_AMT_TR_65"] + "<span                                                                                              \n");
            htmlStr.Append("				style='mso-spacerun: yes'>   </span></td>                                                                                                    \n");
            htmlStr.Append("		                                                                                                                                                     \n");
            htmlStr.Append("			<td class=xl6830999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("		</tr>                                                                                                                                                \n");
            htmlStr.Append("		<tr height=10 style='mso-height-source: userset; height: 2.5pt'>                                                                                     \n");
            htmlStr.Append("			<td height=10 class=xl6730999 style='height: 2.5pt'>&nbsp;</td>                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7330999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7330999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7330999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7330999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7330999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7330999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7330999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7330999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7330999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6830999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("		</tr>                                                                                                                                                \n");
            htmlStr.Append("		<tr class=xl7430999 height=18 style='height: 13.2pt'>                                                                                                \n");
            htmlStr.Append("			<td height=18 class=xl8430999 style='height: 13.2pt'>&nbsp;</td>                                                                                 \n");
            htmlStr.Append("			<td class=xl7430999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td colspan=3 class=xl7930999>Ng&#432;&#7901;i l&#7853;p</td>                                                                                    \n");
            htmlStr.Append("			<td colspan=3 class=xl7930999>Th&#7911; kho xu&#7845;t</td>                                                                                      \n");
            htmlStr.Append("			<td colspan=3 class=xl7930999>Ng&#432;&#7901;i v&#7853;n                                                                                         \n");
            htmlStr.Append("				chuy&#7875;n</td>                                                                                                                            \n");
            htmlStr.Append("			<td colspan=3 class=xl7930999>Th&#7911; kho nh&#7853;p</td>                                                                                      \n");
            htmlStr.Append("			<td class=xl7430999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl8530999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("		</tr>                                                                                                                                                \n");
            htmlStr.Append("		<tr class=xl7430999 height=22                                                                                                                        \n");
            htmlStr.Append("			style='mso-height-source: userset; height: 20.63pt'>                                                                                              \n");
            htmlStr.Append("			<td height=22 class=xl8430999 style='height: 20.63pt'>&nbsp;</td>                                                                                 \n");
            htmlStr.Append("			<td class=xl7430999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td colspan=3 class=xl11630999>Prepared by</td>                                                                                                  \n");
            htmlStr.Append("			<td colspan=3 class=xl11630999>Output WH keeper</td>                                                                                             \n");
            htmlStr.Append("			<td colspan=3 class=xl11630999>Transporter</td>                                                                                                  \n");
            htmlStr.Append("			<td colspan=3 class=xl11630999>Input WH keeper</td>                                                                                              \n");
            htmlStr.Append("			<td class=xl7430999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl8530999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("		</tr>                                                                                                                                                \n");
            htmlStr.Append("		<tr class=xl6930999 height=18                                                                                                                        \n");
            htmlStr.Append("			style='mso-height-source: userset; height: 17.63pt'>                                                                                              \n");
            htmlStr.Append("			<td height=18 class=xl8630999 style='height: 17.63pt'>&nbsp;</td>                                                                                 \n");
            htmlStr.Append("			<td class=xl6930999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td colspan=3 class=xl13130999>(Ký, ghi rõ h&#7885; tên)</td>                                                                                    \n");
            htmlStr.Append("			<td colspan=3 class=xl13130999>(Ký, ghi rõ h&#7885; tên)</td>                                                                                    \n");
            htmlStr.Append("			<td colspan=3 class=xl13130999>(Ký, ghi rõ h&#7885; tên)</td>                                                                                    \n");
            htmlStr.Append("			<td colspan=3 class=xl13130999>(Ký, ghi rõ h&#7885; tên)</td>                                                                                    \n");
            htmlStr.Append("			<td class=xl6930999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl8730999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("		</tr>                                                                                                                                                \n");
            htmlStr.Append("		<tr height=18 style='height: 3.8pt'>                                                                                                                 \n");
            htmlStr.Append("			<td height=18 class=xl6730999 style='height: 3.8pt'>&nbsp;</td>                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7330999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7330999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl9730999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7330999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl9830999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7330999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7330999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7330999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7330999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7330999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7330999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7330999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6830999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("		</tr>                                                                                                                                                \n");
            htmlStr.Append("		<tr height=18 style='height: 3.8pt'>                                                                                                                 \n");
            htmlStr.Append("			<td height=18 class=xl6730999 style='height: 3.8pt'>&nbsp;</td>                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7330999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7330999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7330999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7330999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7330999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7330999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7330999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7330999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7330999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7330999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7330999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7330999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6830999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("		</tr>                                                                                                                                                \n");
            htmlStr.Append("		<tr height=20 style='mso-height-source: userset; height: 5.0pt'>                                                                                     \n");
            htmlStr.Append("			<td height=20 class=xl6730999 style='height: 5.0pt'>&nbsp;</td>                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6830999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("		</tr>                                                                                                                                                \n");
            htmlStr.Append("		<tr height=20 style='mso-height-source: userset; height: 15.0pt'>                                                                                    \n");
            htmlStr.Append("			<td height=20 class=xl6730999 style='height: 15.0pt'>&nbsp;</td>                                                                                 \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("			<td class=xl6630999 colspan=2>&nbsp;</td>                                                                                                        \n");
            htmlStr.Append("			<td colspan=6 class=xl13230999 width=373                                                                                                         \n");
            htmlStr.Append("				style='border-right: 1.0pt solid #0070C0; width: 278pt'>Signature                                                                             \n");
            htmlStr.Append("				Valid<span style='mso-spacerun: yes'> </span>                                                                                                \n");
            htmlStr.Append("			</td>                                                                                                                                            \n");
           
            
            if (dt.Rows[0]["sign_yn"].ToString() == "Y")
            {
                htmlStr.Append("		<td align=left valign=top><![if !vml]><span																											\n");
                htmlStr.Append("			style='mso-ignore: vglayout; position: absolute; z-index: 1; margin-left: -208px; margin-top: 0px; width: 82px; height: 56px'><img              \n");
                htmlStr.Append("				width=82 height=56                                                                                                                          \n");
                htmlStr.Append("				src='${ pageContext.request.contextPath}/ assets / img / check_signed.png'                                                                        \n");
                htmlStr.Append("				v:shapes='Picture_x0020_8 _x0000_s7236'></span> <![endif]><span                                                                             \n");
                htmlStr.Append("			style='mso-ignore: vglayout2'>                                                                                                                  \n");
                htmlStr.Append("				<table cellpadding=0 cellspacing=0>                                                                                                         \n");
                htmlStr.Append("					<tr>                                                                                                                                    \n");
                htmlStr.Append("						<td class=xl6630999>&nbsp;</td>                                                                                                     \n");
                htmlStr.Append("					</tr>                                                                                                                                   \n");
                htmlStr.Append("				</table>                                                                                                                                    \n");
                htmlStr.Append("		</span></td>                                                                                                                                        \n");
            }
            else
            {
                htmlStr.Append("							<td class=xl6630999>&nbsp;</td>                                                                                                                                                                                                    \n");
            }

            htmlStr.Append("			<td class=xl6830999>&nbsp;</td>																						 \n");
            htmlStr.Append("		</tr>                                                                                                                    \n");
            htmlStr.Append("		<tr height=18 style='height: 17.25pt'>                                                                                    \n");
            htmlStr.Append("			<td height=18 class=xl6730999 style='height: 17.25pt'>&nbsp;</td>                                                     \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                      \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                      \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                      \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                      \n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                      \n");
            htmlStr.Append("			<td class=xl6630999 colspan=2>&nbsp;</td>                                                                            \n");

            if (dt.Rows[0]["sign_yn"].ToString() == "Y")
            {
                htmlStr.Append("			<td colspan=6 class=xl13530999 width=373																			\n");
                htmlStr.Append("				style='border-right: 1.0pt solid #0070C0; width: 278pt'>&#272;&#432;&#7907;c                                     \n");
                htmlStr.Append("				ký b&#7903;i:" + dt.Rows[0]["SignedBy"] + "</td>                                                                            \n");
            }
            else
            {
                htmlStr.Append("		<td colspan=6 class=xl13530999 width=373																			\n");
                htmlStr.Append("			style='border-right: 1.0pt solid #0070C0; width: 278pt'>&#272;&#432;&#7907;c                                     \n");
                htmlStr.Append("			ký b&#7903;i:</td>                                                                                              \n");
            }



            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>																								\n");
            htmlStr.Append("			<td class=xl6830999>&nbsp;</td>                                                                                         	\n");
            htmlStr.Append("		</tr>                                                                                                                       	\n");
            htmlStr.Append("		<tr height=18 style='height: 17.25pt'>                                                                                       	\n");
            htmlStr.Append("			<td height=18 class=xl6730999 style='height: 17.25pt'>&nbsp;</td>                                                        	\n");
            htmlStr.Append("			<td class=xl6630999>Mã CQT: " + dt.Rows[0]["cqt_mccqt_id"] + "</td>                                                                           	\n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                         	\n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                         	\n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                         	\n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                         	\n");
            htmlStr.Append("			<td class=xl6630999 colspan=2>&nbsp;</td>                                                                               	\n");
            htmlStr.Append("			<td colspan=6 class=xl12630999                                                                                          	\n");
            htmlStr.Append("				style='border-right: 1.0pt solid #0070C0'><font                                                                      	\n");
            htmlStr.Append("				class='font1430999'>Ngày ký</font><font class='font1330999'>:<span                                                  	\n");
            htmlStr.Append("					style='mso-spacerun: yes'>" + dt.Rows[0]["SignedDate"] + " </span></font></td>                                                    	\n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                         	\n");
            htmlStr.Append("			<td class=xl6830999>&nbsp;</td>                                                                                         	\n");
            htmlStr.Append("		</tr>                                                                                                                       	\n");
            htmlStr.Append("		<tr height=12 style='mso-height-source: userset; height: 9.0pt'>                                                            	\n");
            htmlStr.Append("			<td height=12 class=xl6730999 style='height: 9.0pt'>&nbsp;</td>                                                         	\n");
            htmlStr.Append("			<td class=xl6630999>Tra c&#7913;u t&#7841;i Website: <font                                                              	\n");
            htmlStr.Append("				class='font65513'><span style='mso-spacerun: yes'> </span></font><font                                              	\n");
            htmlStr.Append("				class='font185513'>" + dt.Rows[0]["WEBSITE_EI"] + "</font></td>                                                 	\n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                         	\n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                         	\n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                         	\n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                         	\n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                         	\n");
            htmlStr.Append("			<td class=xl8830999>&nbsp;</td>                                                                                         	\n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;Mã nh&#7853;n hóa &#273;&#417;n: " + dt.Rows[0]["matracuu"] + "</td>                                        	\n");
            htmlStr.Append("			<td class=xl8930999>&nbsp;</td>                                                                                         	\n");
            htmlStr.Append("			<td class=xl8930999>&nbsp;</td>                                                                                         	\n");
            htmlStr.Append("			<td class=xl8930999>&nbsp;</td>                                                                                         	\n");
            htmlStr.Append("			<td class=xl8930999>&nbsp;</td>                                                                                         	\n");
            htmlStr.Append("			<td class=xl8930999>&nbsp;</td>                                                                                         	\n");
            htmlStr.Append("			<td class=xl6630999>&nbsp;</td>                                                                                         	\n");
            htmlStr.Append("			<td class=xl6830999>&nbsp;</td>                                                                                         	\n");
            htmlStr.Append("		</tr>                                                                                                                       	\n");
            htmlStr.Append("		<tr height=18 style='height: 17.25pt'>                                                                                       	\n");
            htmlStr.Append("			<td height=18 class=xl9030999 style='height: 17.25pt'>&nbsp;</td>                                                        	\n");
            htmlStr.Append("			<td class=xl9130999>&nbsp;</td>                                                                                         	\n");
            htmlStr.Append("			<td colspan=12 class=xl12930999>(C&#7847;n ki&#7875;m tra                                                               	\n");
            htmlStr.Append("				&#273;&#7889;i chi&#7871;u khi l&#7853;p, giao, nh&#7853;n hóa                                                      	\n");
            htmlStr.Append("				&#273;&#417;n)</td>                                                                                                 	\n");
            htmlStr.Append("			<td class=xl9130999>&nbsp;</td>                                                                                         	\n");
            htmlStr.Append("			<td class=xl9230999>&nbsp;</td>                                                                                         	\n");
            htmlStr.Append("		</tr>                                                                                                                       	\n");
            htmlStr.Append("		<tr height=18 style='height: 17.25pt'>                                                                                       	\n");
            htmlStr.Append("			<td colspan=16 height=18 class=xl13030999 style='height: 17.25pt'>" + dt.Rows[0]["CONTRACT_INFO_EI"] + "</td>                                                                  	\n");
            htmlStr.Append("		</tr>                                                                                                                       	\n");
            htmlStr.Append("		<![if supportMisalignedColumns]>                                                                                            	\n");
            htmlStr.Append("		<tr height=0 style='display: none'>                                                                                         	\n");
            htmlStr.Append("			<td width=7 style='width: 6.35pt'></td>                                                                                    	\n");
            htmlStr.Append("			<td width=6 style='width: 5.0pt'></td>                                                                                    	\n");
            htmlStr.Append("			<td width=26 style='width: 23.75pt'></td>                                                                                  	\n");
            htmlStr.Append("			<td width=55 style='width: 51.25pt'></td>                                                                                  	\n");
            htmlStr.Append("			<td width=77 style='width: 58pt'></td>                                                                                  	\n");
            htmlStr.Append("			<td width=62 style='width: 58.75pt'></td>                                                                                  	\n");
            htmlStr.Append("			<td width=62 style='width: 58.75pt'></td>                                                                                  	\n");
            htmlStr.Append("			<td width=63 style='width: 58.75pt'></td>                                                                                  	\n");
            htmlStr.Append("			<td width=63 style='width: 58.75pt'></td>                                                                                  	\n");
            htmlStr.Append("			<td width=63 style='width: 58.75pt'></td>                                                                                  	\n");
            htmlStr.Append("			<td width=42 style='width: 31pt'></td>                                                                                  	\n");
            htmlStr.Append("			<td width=42 style='width: 31pt'></td>                                                                                  	\n");
            htmlStr.Append("			<td width=62 style='width: 58.75pt'></td>                                                                                  	\n");
            htmlStr.Append("			<td width=38 style='width: 35.0pt'></td>                                                                                  	\n");
            htmlStr.Append("			<td width=6 style='width: 5.0pt'></td>                                                                                    	\n");
            htmlStr.Append("			<td width=9 style='width: 8.75pt'></td>                                                                                    	\n");
            htmlStr.Append("		</tr>                                                                                                                       	\n");
            htmlStr.Append("		<![endif]>                                                                                                                  	\n"); htmlStr.Append("	</table>                                                                                                                                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                                                                                                                                      \n");
            htmlStr.Append("</body>                                                                                                                                                                                                 \n");
            htmlStr.Append("</html>               \n");

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
            if (ccy != "VND")
            {
                //rtnf += " Đô La Mỹ";
                switch (ccy)
                {
                    case "USD":
                        rtnf += " Đô La Mỹ";
                        break;
                    case "EUR":
                        rtnf += " Euro";
                        break;
                }


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

        private static int countLength(string s)
        {
            int result = 0;
            int max_length = 28;
            int index_length = 0;
            string[] words = s.Split(' ');//tach chuoi dua tren khoang trang
            for (int i = 0; i < words.Length; i++)
            {

                index_length += words[i].Length;
                if (index_length >= max_length)
                {
                    result++;
                    index_length = words[i].Length;

                }
                if (i == words.Length - 1 && result == 0)
                {
                    result = 1;
                }
            }

            return result;
        }


    }
}