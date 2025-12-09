using System;
using System.Collections.Generic;
using System.Data;
//using System.Data.OracleClient;
using System.Drawing;
using System.IO;
using System.Text;
using System.Xml;
using Oracle.DataAccess.Client;
using Oracle.DataAccess.Types;
namespace EInvoice.Company
{
    public class DeaYoung_New_BK20221108
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

                    int count_row = countLength_v2(dt_d.Rows[j]["ITEM_NAME"].ToString());
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
                            page_index[i + 1] = count_row;
                            l_finish = "Y";
                            count_col++;
                            count_col_index++;
                            break;
                        }
                        else
                        {
                            page[i] = count_col_index + 1;
                            page_index[i] = total_countLenght;
                            l_finish = "Y";
                            count_col++;
                            count_col_index++;
                            break;
                        }


                    }
                    else if (total_countLenght > pos_lv)
                    {
                        //if (total_countLenght - pos_lv < 2)
                        //{
                        page[i] = count_col_index + 1;
                        page_index[i] = total_countLenght;
                        count_col++;
                        count_col_index++;
                        break;//continue;	
                        //}
                    }
                    else if (total_countLenght == pos_lv)
                    {
                        page[i] = count_col_index + 1;
                        page_index[i] = total_countLenght;
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
                lb_amount_trans = dt.Rows[0]["EXCHANGERATE"].ToString();
                amount_trans = dt.Rows[0]["TOT_AMT_BK_93"].ToString();
                amount_total = dt.Rows[0]["tot_amt_tr_94"].ToString();
                amount_vat = dt.Rows[0]["VAT_TR_AMT_DIS_TR_91"].ToString();
                amount_net = dt.Rows[0]["NET_TR_AMT_DIS_TR_89"].ToString();

                // read_prive = Num2VNText(dt.Rows[0]["TotalAmountInWord"].ToString(), "USD");
            }

            read_prive = dt.Rows[0]["AMOUNT_WORD_VIE"].ToString();//read_prive.Substring(0, 2) + read_prive.Substring(2, read_prive.Length - 2).ToLower() + '.';
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
            htmlStr.Append("<!--                                                                                                                                                                                                    \n");
            htmlStr.Append("table {                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-displayed-decimal-separator: '\\.';                                                                                                                                                              \n");
            htmlStr.Append("	mso-displayed-thousand-separator: '\\,';                                                                                                                                                             \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".font676511 {                                                                                                                                                                                           \n");
            htmlStr.Append("	color: red;                                                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 14.25pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".font57651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.06pt;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".font67651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.25pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".font77651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".font87651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".font97651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".font107651 {                                                                                                                                                                                           \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.06pt;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".font117651 {                                                                                                                                                                                           \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".font127651 {                                                                                                                                                                                           \n");
            htmlStr.Append("	color: red;                                                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".font137651 {                                                                                                                                                                                           \n");
            htmlStr.Append("	color: #0066CC;                                                                                                                                                                                     \n");
            htmlStr.Append("	font-size: 14.06pt;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".font147651 {                                                                                                                                                                                           \n");
            htmlStr.Append("	color: red;                                                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".font157651 {                                                                                                                                                                                           \n");
            htmlStr.Append("	color: #003366;                                                                                                                                                                                     \n");
            htmlStr.Append("	font-size: 13.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: Calibri, sans-serif;                                                                                                                                                                   \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".font167651 {                                                                                                                                                                                           \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 11.5pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".font177651 {                                                                                                                                                                                           \n");
            htmlStr.Append("	color: red;                                                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 14.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: Calibri, sans-serif;                                                                                                                                                                   \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append(".font17221001 {                                                                                                                                                                                         \n");
            htmlStr.Append("	color: #0070C0;                                                                                                                                                                                     \n");
            htmlStr.Append("	font-size: 11.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: bold;                                                                                                                                                                                   \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append(".font187651 {                                                                                                                                                                                           \n");
            htmlStr.Append("	color: black;                                                                                                                                                                                       \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".font197651 {                                                                                                                                                                                           \n");
            htmlStr.Append("	color: #2F75B5;                                                                                                                                                                                     \n");
            htmlStr.Append("	font-size: 14.25pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".font511323 {                                                                                                                                                                                           \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl657651 {                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 1.06pt;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl657652 {                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.06pt;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl6576511 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 1.0pt;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl667651 {                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 1.0pt;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 0.5pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl677651 {                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 1.0pt;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl687651 {                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 1.0pt;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl697651 {                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl707651 {                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.06pt;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl7076511 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 1.0pt;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl717651 {                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 1.25pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl727651 {                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.06pt;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl737651 {                                                                                                                                                                                             \n");
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
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl747651 {                                                                                                                                                                                             \n");
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
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl757651 {                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.25pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl767652 {                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 1.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: 0.5pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 0.5pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl767653 {                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 1.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");

            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl777651 {                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.06pt;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: '\\@';                                                                                                                                                                            \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl787651 {                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.25pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl797651 {                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.06pt;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl807651 {                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.06pt;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl817651 {                                                                                                                                                                                             \n");
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
            htmlStr.Append("	text-align: right;                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl827651 {                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 1.06pt;                                                                                                                                                                                   \n");
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
            htmlStr.Append("	border-bottom: 0.5pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: 0.5pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl837651 {                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.25pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl847651 {                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.06pt;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl857651 {                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl867651 {                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.25pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 0.5pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl877651 {                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.25pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: '\\@';                                                                                                                                                                            \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl887651 {                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl897651 {                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 1.06pt;                                                                                                                                                                                 \n");
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
            htmlStr.Append("	border-bottom: 0.5pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl907651 {                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 1.06pt;                                                                                                                                                                                 \n");
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
            htmlStr.Append("	border-bottom: 0.5pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl917651 {                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 1.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: 0.5pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl927651 {                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.25pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl937651 {                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.25pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl947651 {                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.06pt;                                                                                                                                                                                   \n");
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
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 0.5pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl957651 {                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
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
            htmlStr.Append("	border-left: 0.5pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl967651 {                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl977651 {                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.06pt;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 0.5pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl987651 {                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.25pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl997651 {                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.06pt;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1007651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 1.25pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1017651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: #C00000;                                                                                                                                                                                     \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1027651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.25pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: 0.5pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1037651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 0.5pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: 0.5pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1047651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 17.1pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1057651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 17.1pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1067651 {                                                                                                                                                                                            \n");
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
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1077651 {                                                                                                                                                                                            \n");
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
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1087651 {                                                                                                                                                                                            \n");
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
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1097651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: red;                                                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 14.06pt;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1107651 {                                                                                                                                                                                            \n");
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
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1117651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: red;                                                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 14.06pt;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 0.5pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1127651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.25pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1137651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1147651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: #0070C0;                                                                                                                                                                                     \n");
            htmlStr.Append("	font-size: 14.06pt;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1157651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.06pt;                                                                                                                                                                                   \n");
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
            htmlStr.Append("	border-left: 0.5pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1167651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.06pt;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1177651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.25pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1187651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: red;                                                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1197651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 17.1pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 0.5pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1207651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.06pt;                                                                                                                                                                                   \n");
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
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1217651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: '\\@';                                                                                                                                                                            \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1.0pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: 1.0pt dotted windowtext;                                                                                                                                                              \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1227651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: '\\@';                                                                                                                                                                            \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1.0pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: 1.0pt dotted windowtext;                                                                                                                                                              \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1237651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: '\\@';                                                                                                                                                                            \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1.0pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1247651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.25pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: '\\@';                                                                                                                                                                            \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: 0.5pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1257651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1.0pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: 1.0pt dotted windowtext;                                                                                                                                                              \n");
            htmlStr.Append("	border-left: 0.5pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1267651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: '\\@';                                                                                                                                                                            \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1.0pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: 1.0pt dotted windowtext;                                                                                                                                                              \n");
            htmlStr.Append("	border-left: 0.5pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1277651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1.0pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1287651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: '\\@';                                                                                                                                                                            \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1.0pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 0.5pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1297651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: '\\@';                                                                                                                                                                            \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1.0pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 0.5pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1307651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: '\\@';                                                                                                                                                                            \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1.0pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1317651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: '\\@';                                                                                                                                                                            \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1.0pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1327651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 0.5pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1337651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: '\\@';                                                                                                                                                                            \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1.0pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1347651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: red;                                                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 14.06pt;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1357651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: 0%;                                                                                                                                                                              \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 0.5pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1367651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.25pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1377651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 1.06pt;                                                                                                                                                                                   \n");
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
            htmlStr.Append("	border-left: 0.5pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl13776511 {                                                                                                                                                                                           \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 1.0pt;                                                                                                                                                                                   \n");
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
            htmlStr.Append("	border-left: 0.5pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1387651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1397651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 0.5pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1407651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: '\\@';                                                                                                                                                                            \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1.0pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 1.0pt dotted windowtext;                                                                                                                                                              \n");
            htmlStr.Append("	border-left: 0.5pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1417651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: '\\@';                                                                                                                                                                            \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1.0pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 1.0pt dotted windowtext;                                                                                                                                                              \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1427651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 0.5pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1437651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.25pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1447651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: red;                                                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1457651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: black;                                                                                                                                                                                       \n");
            htmlStr.Append("	font-size: 14.06pt;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 0.5pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: 0.5pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1467651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: black;                                                                                                                                                                                       \n");
            htmlStr.Append("	font-size: 14.06pt;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 0.5pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1477651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: black;                                                                                                                                                                                       \n");
            htmlStr.Append("	font-size: 14.06pt;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 0.5pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1487651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: red;                                                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 20.20pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1497651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: red;                                                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 14.0pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1507651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.25pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1517651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.25pt;                                                                                                                                                                                  \n");
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
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1527651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1537651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1547651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
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
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1557651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
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
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1567651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.25pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 0.5pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1577651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.25pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1587651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.06pt;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 0.5pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1597651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.06pt;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 0.5pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: 0.5pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1607651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.06pt;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 0.5pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1617651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.06pt;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: 0.5pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1627651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.06pt;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1637651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 0.5pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: 0.5pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1647651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: 0.5pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1657651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1.0pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 1.0pt dotted windowtext;                                                                                                                                                              \n");
            htmlStr.Append("	border-left: 0.5pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1667651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1.0pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: 1.0pt dotted windowtext;                                                                                                                                                              \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1677651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1.0pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 1.0pt dotted windowtext;                                                                                                                                                              \n");
            htmlStr.Append("	border-left: 0.5pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1687651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.25pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1.0pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 1.0pt dotted windowtext;                                                                                                                                                              \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1697651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.25pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1.0pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: 1.0pt dotted windowtext;                                                                                                                                                              \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1707651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: '\\@';                                                                                                                                                                            \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 1.0pt dotted windowtext;                                                                                                                                                              \n");
            htmlStr.Append("	border-left: 0.5pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1717651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: '\\@';                                                                                                                                                                            \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 1.0pt dotted windowtext;                                                                                                                                                              \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1727651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: '\\#\\,\\#\\#0';                                                                                                                                                                     \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1.0pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 1.0pt dotted windowtext;                                                                                                                                                              \n");
            htmlStr.Append("	border-left: 0.5pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1737651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: '\\#\\,\\#\\#0';                                                                                                                                                                     \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1.0pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 1.0pt dotted windowtext;                                                                                                                                                              \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1747651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: red;                                                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 14.25pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1.0pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 1.0pt dotted windowtext;                                                                                                                                                              \n");
            htmlStr.Append("	border-left: 0.5pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1757651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: red;                                                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 14.25pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1.0pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 1.0pt dotted windowtext;                                                                                                                                                              \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1767651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: red;                                                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 14.25pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1.0pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: 1.0pt dotted windowtext;                                                                                                                                                              \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1777651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1.0pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 0.5pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append(".xl1777652 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 1.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1.0pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 0.5pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1787651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1.0pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1797651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 12pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1.0pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1807651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: '\\@';                                                                                                                                                                            \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1.0pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 0.5pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1817651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: '\\@';                                                                                                                                                                            \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 1.0pt dotted windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1827651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1837651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.25pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1847651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 0.5pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1857651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: 0.5pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1867651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: '\\@';                                                                                                                                                                            \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 0.5pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: 0.5pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1877651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: '\\@';                                                                                                                                                                            \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 0.5pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1887651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: black;                                                                                                                                                                                       \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 0.5pt solid windowtext;                                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1897651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: black;                                                                                                                                                                                       \n");
            htmlStr.Append("	font-size: 13.06pt;                                                                                                                                                                                 \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1907651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: red;                                                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 14.06pt;                                                                                                                                                                                   \n");
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
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1917651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: red;                                                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 14.06pt;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1927651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: red;                                                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 14.06pt;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: 'Short Date';                                                                                                                                                                    \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 0.5pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1937651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: red;                                                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 14.06pt;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: 'Short Date';                                                                                                                                                                    \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: 0.5pt solid windowtext;                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: 0.5pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1947651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 10.06pt;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 0.5pt solid windowtext;                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1957651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 10.06pt;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1967651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.06pt;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1977651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.25pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl1987651 {                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size: 14.25pt;                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                   \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                 \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                              \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                         \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                             \n");
            htmlStr.Append("	border-top: 0.5pt solid windowtext;                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                  \n");
            htmlStr.Append("	background: white;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                                            \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                       \n");
            htmlStr.Append("-->                                                                                                                                                                                                     \n");
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

            double v_totalHeightLastPage = 153.5; // 173.5; //203.5 243.5.5;

            double v_totalHeightPage = 570;//   540;

            for (int k = 0; k < v_countNumberOfPages; k++)
            {
                v_totalHeightPage = 545;  //492;//

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

                htmlStr.Append("	<table border=0 cellpadding=0 cellspacing=0 width=742 class=xl657651                                                                                                                                \n");
                htmlStr.Append("		style='border-collapse: collapse; table-layout: fixed; width: 657pt'>                                                                                                                         \n");
                htmlStr.Append("		<col class=xl657651 width=6                                                                                                                                                                     \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 199; width: 4.75pt'>                                                                                                                        \n");
                htmlStr.Append("		<col class=xl657651 width=33                                                                                                                                                                    \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 1166; width: 29.68pt'>                                                                                                                     \n");
                htmlStr.Append("		<col class=xl657651 width=70                                                                                                                                                                    \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 2474; width: 118pt'>                                                                                                                      \n");
                htmlStr.Append("		<col class=xl657651 width=55                                                                                                                                                                    \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 1962; width: 33.69pt'>                                                                                                                     \n");
                htmlStr.Append("		<col class=xl657651 width=41                                                                                                                                                                    \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 1450; width: 36.81pt'>                                                                                                                     \n");
                htmlStr.Append("		<col class=xl657651 width=98                                                                                                                                                                    \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 3498; width: 87.88pt'>                                                                                                                      \n");
                htmlStr.Append("		<col class=xl657651 width=27                                                                                                                                                                    \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 967; width: 23.75pt'>                                                                                                                         \n");
                htmlStr.Append("		<col class=xl657651 width=78                                                                                                                                                                    \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 2759; width: 68.88pt'>                                                                                                                      \n");
                htmlStr.Append("		<col class=xl657651 width=56                                                                                                                                                                    \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 1991; width: 49.88pt'>                                                                                                                      \n");
                htmlStr.Append("		<col class=xl657651 width=30                                                                                                                                                                    \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 1080; width: 27.32pt'>                                                                                                                     \n");
                htmlStr.Append("		<col class=xl657651 width=6                                                                                                                                                                     \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 199; width: 4.75pt'>                                                                                                                        \n");
                htmlStr.Append("		<col class=xl657651 width=49                                                                                                                                                                    \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 1735; width: 33.94pt'>                                                                                                                     \n");
                htmlStr.Append("		<col class=xl657651 width=48                                                                                                                                                                    \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 1706; width: 42.75pt'>                                                                                                                      \n");
                htmlStr.Append("		<col class=xl657651 width=6                                                                                                                                                                     \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 199; width: 4.75pt'>                                                                                                                        \n");
                htmlStr.Append("		<col class=xl657651 width=55                                                                                                                                                                    \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 1962; width: 48.68pt'>                                                                                                                     \n");
                htmlStr.Append("		<col class=xl657651 width=78                                                                                                                                                                    \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 2759; width: 68.88pt'>                                                                                                                      \n");
                htmlStr.Append("		<col class=xl657651 width=6                                                                                                                                                                     \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 199; width: 4.75pt'>                                                                                                                        \n");
                htmlStr.Append("		<tr height=33 style='mso-height-source: userset; height: 29.75pt'>                                                                                                                               \n");
                htmlStr.Append("			<td height=33 class=xl1197651 width=6 style='height: 29.75pt; width: 4.75pt'                                                                                                                                  \n");
                htmlStr.Append("				align=left valign=top><![if !vml]><span                                                                                                                                                 \n");
                htmlStr.Append("				style='mso-ignore: vglayout; position: absolute; z-index: 2; margin-left: 10px; margin-top: 6px; width: 206px; height: 40px'><img                                                   \n");
                htmlStr.Append("					width=206 height=40                                                                                                                                                            \n");
                htmlStr.Append("					src='D:\\webproject\\e-invoice-ws\\02.Web\\EInvoice\\img\\DAEYOUNG_002_NEW.png'                                                                                                         \n");
                htmlStr.Append("					alt='Description: logo4' v:shapes='Picture_x0020_1'></span>                                                                                                                         \n");
                htmlStr.Append("			<![endif]><span style='mso-ignore: vglayout2'>                                                                                                                                              \n");
                htmlStr.Append("					<table cellpadding=0 cellspacing=0>                                                                                                                                                 \n");
                htmlStr.Append("						<tr>                                                                                                                                                                            \n");
                htmlStr.Append("							<td height=33  width=6                                                                                                                                       \n");
                htmlStr.Append("								style='height: 29.75pt; width: 4.75pt'>&nbsp;</td>                                                                                                                        \n");
                htmlStr.Append("						</tr>                                                                                                                                                                           \n");
                htmlStr.Append("					</table>                                                                                                                                                                            \n");
                htmlStr.Append("			</span></td>                                                                                                                                                                                \n");
                htmlStr.Append("			<td class=xl677651 width=33 style='width: 29.68pt'>&nbsp;</td>                                                                                                                              \n");
                htmlStr.Append("			<td class=xl1047651 width=70 style='width: 49.4pt'>&nbsp;</td>                                                                                                                              \n");
                htmlStr.Append("			<td class=xl1047651 width=55 style='width: 48.68pt'>&nbsp;</td>                                                                                                                             \n");
                htmlStr.Append("			<td colspan=8 class=xl1487651 width=385 style='width: 289pt'>HÓA                                                                                                                            \n");
                htmlStr.Append("				&#272;&#416;N GIÁ TR&#7882; GIA T&#258;NG</td>                                                                                                                                          \n");
                htmlStr.Append("			<td class=xl1047651 width=48 style='width: 42.75pt'>&nbsp;</td>                                                                                                                              \n");
                htmlStr.Append("			<td class=xl1047651 width=6 style='width: 4.75pt'>&nbsp;</td>                                                                                                                                \n");
                htmlStr.Append("			<td class=xl1047651 width=55 style='width: 48.68pt'>&nbsp;</td>                                                                                                                             \n");
                htmlStr.Append("			<td class=xl1047651 width=78 style='width: 68.88pt'>&nbsp;</td>                                                                                                                              \n");
                htmlStr.Append("			<td class=xl1057651 width=6 style='width: 4.75pt'>&nbsp;</td>                                                                                                                                \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr height=24 style='mso-height-source: userset; height: 24.0pt'>                                                                                                                               \n");
                htmlStr.Append("			<td height=24 class=xl1377651 style='height: 24.0pt'>&nbsp;</td>                                                                                                                            \n");
                htmlStr.Append("			<td class=xl657651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("			<td class=xl657651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("			<td class=xl657651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("			<td colspan=8 class=xl1497651>(VAT INVOICE)</td>                                                                                                                                            \n");
                htmlStr.Append("			<td class=xl697651 colspan=4></td>                                                                                                                                          \n");
                htmlStr.Append("			<td class=xl707651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr height=21 style='mso-height-source: userset; height: 20.06pt'>                                                                                                                              \n");
                htmlStr.Append("			<td height=21 class=xl1377651 style='height: 20.06pt'>&nbsp;</td>                                                                                                                           \n");
                htmlStr.Append("			<td class=xl657651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("			<td class=xl657651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("			<td class=xl657651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("			<td class=xl657651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("			<td class=xl657651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("			<td class=xl657651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("			<td class=xl657651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("			<td class=xl657651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("			<td class=xl657651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("			<td class=xl657651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("			<td class=xl657651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("			<td class=xl697651 colspan=4>Ký hi&#7879;u (<font                                                                                                                                           \n");
                htmlStr.Append("				class='font107651'>Serial</font><font class='font87651'>):</font><font                                                                                                                  \n");
                htmlStr.Append("				class='font77651'> " + dt.Rows[0]["templateCode"] + "" + dt.Rows[0]["InvoiceSerialNo"] + "</font></td>                                                                                                                                        \n");
                htmlStr.Append("			<td class=xl707651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr height=24 style='mso-height-source: userset; height: 24.0pt'>                                                                                                                               \n");
                htmlStr.Append("			<td height=24 class=xl1377651 style='height: 24.0pt'>&nbsp;</td>                                                                                                                            \n");
                htmlStr.Append("			<td class=xl657651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("			<td class=xl657651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("			<td class=xl657651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("			<td colspan=8 class=xl1507651 width=385 style='width:289pt'>Ngày <font                                                                                                                      \n");
                htmlStr.Append("				  class='font107651'>(Date)</font><font class='font67651'>&nbsp;" + dt.Rows[0]["invoiceissueddate_dd"] + "&nbsp;<span                                                                                                \n");
                htmlStr.Append("				  style='mso-spacerun:yes'>   </span></font><font class='font67651'><span                                                                                                               \n");
                htmlStr.Append("				  style='mso-spacerun:yes'> </span>tháng </font><font class='font107651'>(month)</font><font                                                                                            \n");
                htmlStr.Append("				  class='font67651'>&nbsp;" + dt.Rows[0]["invoiceissueddate_mm"] + "&nbsp;<span style='mso-spacerun:yes'>    </span>n&#259;m </font><font                                                                            \n");
                htmlStr.Append("				  class='font107651'>(year)</font><font class='font107651'>&nbsp;" + dt.Rows[0]["invoiceissueddate_yyyy"] + "&nbsp;</br>" + v_titlePageNumber + "<span                                                                                             \n");
                htmlStr.Append("				  style='mso-spacerun:yes'>    </span></font>                                                                                                                                           \n");
                htmlStr.Append("			 </td>                                                                                                                                                                                      \n");
                htmlStr.Append("			<td class=xl697651 colspan=3>S&#7889; (<font class='font107651'>No</font><font                                                                                                              \n");
                htmlStr.Append("				class='font87651'>.):</font><font class='font77651'><span                                                                                                                               \n");
                htmlStr.Append("					style='mso-spacerun: yes'>      </span></font><font class='font147651'><span                                                                                                        \n");
                htmlStr.Append("					style='mso-spacerun: yes'> </span></font><font class='font177651'>" + dt.Rows[0]["InvoiceNumber"] + "</font></td>                                                                                     \n");
                htmlStr.Append("			<td class=xl657651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("			<td class=xl727651 width=6 style='width: 4.75pt'>&nbsp;</td>                                                                                                                                 \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr height=4 style='mso-height-source: userset; height: 3.0pt'>                                                                                                                                 \n");
                htmlStr.Append("			<td colspan=2 height=4 class=xl13776511 style='height: 3.0pt'>&nbsp;</td>                                                                                                                   \n");
                htmlStr.Append("			<td class=xl6576511 colspan=14>&nbsp;</td>                                                                                                                                                  \n");
                htmlStr.Append("			<td class=xl7076511>&nbsp;</td>                                                                                                                                                             \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr height=4 style='mso-height-source: userset; height: 3.0pt'>                                                                                                                                 \n");
                htmlStr.Append("			<td height=4 class=xl667651 style='height: 3.0pt'>&nbsp;</td>                                                                                                                               \n");
                htmlStr.Append("			<td class=xl677651 colspan=15>&nbsp;</td>                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl687651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr height=24 style='mso-height-source: userset; height: 24.0pt'>                                                                                                                               \n");
                htmlStr.Append("			<td height=24 class=xl1377651 style='height: 24.0pt'>&nbsp;</td>                                                                                                                            \n");
                htmlStr.Append("			<td class=xl757651 colspan=12>&#272;&#417;n v&#7883; bán hàng (<font                                                                                                                        \n");
                htmlStr.Append("				class='font107651'>Seller</font><font class='font67651'>): </font><font                                                                                                                 \n");
                htmlStr.Append("				class='font197651'>" + dt.Rows[0]["Seller_Name"] + "</font></td>                                                                                                                                       \n");
                htmlStr.Append("			<td class=xl657651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("			<td colspan=2 class=xl1967651>&nbsp;</td>                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl707651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr height=24 style='mso-height-source: userset; height: 24.0pt'>                                                                                                                               \n");
                htmlStr.Append("			<td height=24 class=xl1377651 style='height: 24.0pt'>&nbsp;</td>                                                                                                                            \n");
                htmlStr.Append("			<td class=xl757651 colspan=5>Mã s&#7889; thu&#7871; (<font                                                                                                                                  \n");
                htmlStr.Append("				class='font107651'>Tax code</font><font class='font67651'>):<font                                                                                                                       \n");
                htmlStr.Append("					class='font117651'> </font><font class='font157651'>" + dt.Rows[0]["Seller_TaxCode"] + "</font></td>                                                                                             \n");
                htmlStr.Append("			<td class=xl767651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("			<td class=xl757651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("			<td class=xl757651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("			<td class=xl757651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("			<td class=xl657651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("			<td class=xl657651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("			<td class=xl657651 colspan=3>&nbsp;</td>                                                                                                                                                    \n");
                htmlStr.Append("			<td class=xl657651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("			<td class=xl707651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr height=24 style='mso-height-source: userset; height: 24.0pt'>                                                                                                                               \n");
                htmlStr.Append("			<td height=24 class=xl1377651 style='height: 24.0pt'>&nbsp;</td>                                                                                                                            \n");
                htmlStr.Append("			<td class=xl757651 colspan=2>&#272;&#7883;a ch&#7881; (<font                                                                                                                                \n");
                htmlStr.Append("				class='font107651'>Address</font><font class='font67651'>):                                                                                                                             \n");
                htmlStr.Append("			</font></td>                                                                                                                                                                                \n");
                htmlStr.Append("			<td class=xl757651 colspan=13>" + dt.Rows[0]["SELLER_ADDRESS"] + "</td>                                                                                                                                  \n");
                htmlStr.Append("			<td class=xl707651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr height=24 style='mso-height-source: userset; height: 24.0pt'>                                                                                                                               \n");
                htmlStr.Append("			<td height=24 class=xl1377651 style='height: 24.0pt'>&nbsp;</td>                                                                                                                            \n");
                htmlStr.Append("			<td class=xl757651 colspan=5>&#272;i&#7879;n tho&#7841;i (<font                                                                                                                             \n");
                htmlStr.Append("				class='font107651'>Tel </font><font class='font67651'>): " + dt.Rows[0]["Seller_Tel"] + "</td>                                                                                                       \n");
                htmlStr.Append("			<td class=xl777651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("			<td class=xl757651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("			<td class=xl757651 colspan=4>Fax: " + dt.Rows[0]["Seller_Fax"] + "</td>                                                                                                                                  \n");
                htmlStr.Append("			<td class=xl657651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("			<td class=xl657651 colspan=3>&nbsp;</td>                                                                                                                                                    \n");
                htmlStr.Append("			<td class=xl707651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr height=24 style='mso-height-source: userset; height: 24.0pt'>                                                                                                                               \n");
                htmlStr.Append("			<td height=24 class=xl1377651 style='height: 24.0pt'>&nbsp;</td>                                                                                                                            \n");
                htmlStr.Append("			<td class=xl757651 colspan=3>S&#7889; tài kho&#7843;n (<font                                                                                                                                \n");
                htmlStr.Append("				class='font107651'>Acc. code</font><font class='font67651'>)</font><span                                                                                                                \n");
                htmlStr.Append("				style='display: none'><font class='font67651'>:<span                                                                                                                                    \n");
                htmlStr.Append("						style='mso-spacerun: yes'> </span></font></span></td>                                                                                                                           \n");
                htmlStr.Append("			<td colspan=12 class=xl1517651 width=572 style='width: 428pt'>" + dt.Rows[0]["SELLER_ACCOUNTNO"] + " " + dt.Rows[0]["BANK_NM78"] + "</td>                                                                                                     \n");
                htmlStr.Append("			<td class=xl707651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr height=24 style='mso-height-source: userset; height: 24.0pt'>                                                                                                                               \n");
                htmlStr.Append("			<td height=24 class=xl1377651 style='height: 24.0pt'>&nbsp;</td>                                                                                                                            \n");
                htmlStr.Append("			<td class=xl757651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("			<td class=xl757651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("			<td class=xl757651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("			<td colspan=12 class=xl1517651 width=572 style='width: 428pt'>" + dt.Rows[0]["SELLER_ACCOUNTNO2"] + " " + dt.Rows[0]["BANK_NM79"] + "</td>                                                                                                   \n");
                htmlStr.Append("			<td class=xl707651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr height=4 style='mso-height-source: userset; height: 3.0pt'>                                                                                                                                 \n");
                htmlStr.Append("			<td colspan=2 height=4 class=xl13776511 style='height: 3.0pt'>&nbsp;</td>                                                                                                                   \n");
                htmlStr.Append("			<td class=xl6576511 colspan=14>&nbsp;</td>                                                                                                                                                  \n");
                htmlStr.Append("			<td class=xl7076511>&nbsp;</td>                                                                                                                                                             \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr height=4 style='mso-height-source: userset; height: 3.0pt'>                                                                                                                                 \n");
                htmlStr.Append("			<td height=4 class=xl667651 style='height: 3.0pt'>&nbsp;</td>                                                                                                                               \n");
                htmlStr.Append("			<td class=xl677651 colspan=15>&nbsp;</td>                                                                                                                                                   \n");
                htmlStr.Append("			<td class=xl687651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr height=24 style='mso-height-source: userset; height: 24.0pt'>                                                                                                                               \n");
                htmlStr.Append("			<td height=24 class=xl1377651 style='height: 24.0pt'>&nbsp;</td>                                                                                                                            \n");
                htmlStr.Append("			<td class=xl757651 colspan=5>H&#7885; tên ng&#432;&#7901;i mua                                                                                                                              \n");
                htmlStr.Append("				hàng (<font class='font107651'>Customer's name</font><font                                                                                                                              \n");
                htmlStr.Append("				class='font67651'>):</font>                                                                                                                                                             \n");
                htmlStr.Append("			</td>                                                                                                                                                                                       \n");
                htmlStr.Append("			<td colspan=11 class=xl1547651                                                                                                                                                              \n");
                htmlStr.Append("				style='border-right: 0.5pt solid black'>&nbsp;" + dt.Rows[0]["buyer"] + "</td>                                                                                                                    \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr height=29 style='mso-height-source: userset; height: 24.0pt'>                                                                                                                              \n");
                htmlStr.Append("			<td height=29 class=xl1377651 style='height: 26.50pt'>&nbsp;</td>                                                                                                                           \n");
                htmlStr.Append("			<td class=xl757651 colspan=4>Tên &#273;&#417;n v&#7883; (<font                                                                                                                              \n");
                htmlStr.Append("				class='font107651'>Company's name</font><font class='font67651'>):</font></td>                                                                                                          \n");
                htmlStr.Append("			<td colspan=11 class=xl1557651>" + dt.Rows[0]["buyerlegalname"] + "</td>                                                                                                                                      \n");
                htmlStr.Append("			<td class=xl1387651>&nbsp;</td>                                                                                                                                                             \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr height=24 style='mso-height-source: userset; height: 24.0pt'>                                                                                                                               \n");
                htmlStr.Append("			<td height=24 class=xl1377651 style='height: 24.0pt'>&nbsp;</td>                                                                                                                            \n");
                htmlStr.Append("			<td class=xl757651 colspan=15>Mã s&#7889; thu&#7871; (<font                                                                                                                                 \n");
                htmlStr.Append("				class='font107651'>Tax code</font><font class='font67651'>):<span                                                                                                                       \n");
                htmlStr.Append("					style='mso-spacerun: yes'> </span>" + dt.Rows[0]["BuyerTaxCode"] + "</font></td>                                                                                                                      \n");
                htmlStr.Append("			<td class=xl1387651>&nbsp;</td>                                                                                                                                                             \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr height=26 style='mso-height-source: userset; height: 24.0pt'>                                                                                                                               \n");
                htmlStr.Append("			<td height=26 class=xl1377651 style='height: 24.0pt'>&nbsp;</td>                                                                                                                            \n");
                htmlStr.Append("			<td class=xl757651 colspan=2>&#272;&#7883;a ch&#7881; (<font                                                                                                                                \n");
                htmlStr.Append("				class='font107651'>Address</font><font class='font67651'>):</font></td>                                                                                                                 \n");
                htmlStr.Append("			<td class=xl757651 colspan=13><font class='font67651'>" + dt.Rows[0]["BuyerAddress"] + "</font></td>                                                                                                           \n");
                htmlStr.Append("			<td class=xl807651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr height=26 style='mso-height-source: userset; height: 24.0pt'>                                                                                                                               \n");
                htmlStr.Append("			<td height=26 class=xl1377651 style='height: 24.0pt'>&nbsp;</td>                                                                                                                            \n");
                htmlStr.Append("			<td class=xl757651 colspan=5>Hình th&#7913;c thanh toán (<font                                                                                                                              \n");
                htmlStr.Append("				class='font107651'>Payment Method)</font><font class='font67651'>:<span                                                                                                                 \n");
                htmlStr.Append("					style='mso-spacerun: yes'> </span>" + dt.Rows[0]["PaymentMethodCK"] + "</font></td>                                                                                                                \n");
                htmlStr.Append("			<td class=xl817651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("			<td class=xl757651></td>                                                                                                                                                              \n");
                htmlStr.Append("			<td class=xl757651 colspan=8>&nbsp;Đơn vị tiền tệ <font class='font107651'>(Currency)</font> : " + dt.Rows[0]["CurrencyCodeUSD"] + "</td>                                                                                                                                                              \n");
                htmlStr.Append("			<td class=xl707651>&nbsp;</td>                                                                                                                                                              \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr height=4 style='mso-height-source: userset; height: 3.0pt'>                                                                                                                                 \n");
                htmlStr.Append("			<td colspan=2 height=4 class=xl13776511 style='height: 3.0pt'>&nbsp;</td>                                                                                                                   \n");
                htmlStr.Append("			<td class=xl6576511 colspan=14>&nbsp;</td>                                                                                                                                                  \n");
                htmlStr.Append("			<td class=xl7076511>&nbsp;</td>                                                                                                                                                             \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr class=xl837651 height=21 style='height: 19.5pt'>                                                                                                                                            \n");
                htmlStr.Append("			<td colspan=2 height=21 class=xl1567651 style='height: 19.5pt'>STT</td>                                                                                                                     \n");
                htmlStr.Append("			<td colspan=5 class=xl1567651 style='border-right: 0.5pt solid black'>Tên                                                                                                                    \n");
                htmlStr.Append("				hàng hóa, d&#7883;ch v&#7909;</td>                                                                                                                                                      \n");
                htmlStr.Append("			<td class=xl1437651>&#272;&#417;n v&#7883; tính</td>                                                                                                                                        \n");
                htmlStr.Append("			<td colspan=3 class=xl1567651 style='border-right: 0.5pt solid black'>S&#7889;                                                                                                               \n");
                htmlStr.Append("				l&#432;&#7907;ng</td>                                                                                                                                                                   \n");
                htmlStr.Append("			<td colspan=3 class=xl1437651>&#272;&#417;n giá</td>                                                                                                                                        \n");
                htmlStr.Append("			<td colspan=3 class=xl1567651 style='border-right: 0.5pt solid black'>Thành                                                                                                                  \n");
                htmlStr.Append("				ti&#7873;n</td>                                                                                                                                                                         \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr class=xl847651 height=18 style='height: 16.5pt'>                                                                                                                                            \n");
                htmlStr.Append("			<td colspan=2 height=18 class=xl1587651 style='height: 16.5pt'>No.</td>                                                                                                                     \n");
                htmlStr.Append("			<td colspan=5 class=xl1597651 style='border-right: 0.5pt solid black'>Name                                                                                                                   \n");
                htmlStr.Append("				of goods and services</td>                                                                                                                                                              \n");
                htmlStr.Append("			<td class=xl847651>Unit</td>                                                                                                                                                                \n");
                htmlStr.Append("			<td colspan=3 class=xl1587651 style='border-right: 0.5pt solid black'>Quantity</td>                                                                                                          \n");
                htmlStr.Append("			<td colspan=3 class=xl847651>Unit price</td>                                                                                                                                                \n");
                htmlStr.Append("			<td colspan=3 class=xl1587651 style='border-right: 0.5pt solid black'>Amount</td>                                                                                                            \n");
                htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
                htmlStr.Append("		<tr class=xl857651 height=20                                                                                                                                                                    \n");
                htmlStr.Append("			style='mso-height-source: userset; height: 18.75pt'>                                                                                                                                         \n");
                htmlStr.Append("			<td colspan=2 height=20 class=xl1637651 style='height: 18.75pt'>A</td>                                                                                                                       \n");
                htmlStr.Append("			<td colspan=5 class=xl1637651 style='border-right: 0.5pt solid black'>                                                                                                                       \n");
                htmlStr.Append("				<![if !vml]><span                                                                                                                                                                       \n");
                htmlStr.Append("				style='mso-ignore: vglayout; position: absolute; z-index: 2; margin-left: 0px; margin-top: 60px; width: 649.35px; height: 52.65px'><img                                              \n");
                htmlStr.Append("					width=811 height=65                                                                                                                                                           \n");
                htmlStr.Append("					src='D:\\webproject\\e-invoice-ws\\02.Web\\EInvoice\\img\\DAEYOUNG_003.png'                                                                                                             \n");
                htmlStr.Append("					alt='Description: logo4' v:shapes='Picture_x0020_1'></span>                                                                                                                         \n");
                htmlStr.Append("			<![endif]><span style='mso-ignore: vglayout2'>                                                                                                                                              \n");
                htmlStr.Append("					<table cellpadding=0 cellspacing=0>                                                                                                                                                 \n");
                htmlStr.Append("						<tr>                                                                                                                                                                            \n");
                htmlStr.Append("							<td colspan=5></td>                                                                                                                                                         \n");
                htmlStr.Append("						</tr>                                                                                                                                                                           \n");
                htmlStr.Append("					</table>                                                                                                                                                                            \n");
                htmlStr.Append("			</span>                                                                                                                                                                                     \n");
                htmlStr.Append("			B</td>                                                                                                                                                                                      \n");
                htmlStr.Append("			<td class=xl1397651>C</td>                                                                                                                                                                  \n");
                htmlStr.Append("			<td colspan=3 class=xl1637651 style='border-right: 0.5pt solid black'>1</td>                                                                                                                 \n");
                htmlStr.Append("			<td colspan=3 class=xl1397651>2</td>                                                                                                                                                        \n");
                htmlStr.Append("			<td colspan=3 class=xl1637651 style='border-right: 0.5pt solid black'>3                                                                                                                      \n");
                htmlStr.Append("				= 1 x 2</td>                                                                                                                                                                            \n");
                htmlStr.Append("		</tr>    \n");


                v_rowHeight = "22.0pt"; //"26.5pt";
                v_rowHeightEmpty = "22.0pt";
                v_rowHeightNumber = 22;

                v_rowHeightLast = "22.0pt";// "23.5pt";
                v_rowHeightLastNumber = 22;// 23.5;
                v_rowHeightEmptyLast = "20.0pt"; //"23.5pt";


                //String[] item_name = null;


                List<string> item_name = null;

                for (int dtR = 0; dtR < page[k]; dtR++) //dt.size()                                                                                                                                             
                {
                    item_name = new List<string>(dt_d.Rows[v_index]["ITEM_NAME"].ToString().Split(new string[] { "&#xA;" }, StringSplitOptions.None));

                    if (dtR == 0)
                    {
                        if (item_name.Count == 1)
                        {

                            if (k == v_countNumberOfPages - 1)
                            {
                                v_totalHeightLastPage = v_totalHeightLastPage - v_rowHeightLastNumber;
                            }
                            else
                            {
                                v_totalHeightPage = v_totalHeightPage - (v_rowHeightNumber);
                            }

                            htmlStr.Append("				<tr class=xl837651 height=25                                                                                                                                                             \n");
                            htmlStr.Append("					style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                                                                                 \n");
                            htmlStr.Append("					<td colspan=2 height=25 class=xl1657651 width=39                                                                                                                                     \n");
                            htmlStr.Append("						style='border-right: .5pt solid black; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "; width: 29pt'>&nbsp;" + dt_d.Rows[v_index][7] + "</td>                                                                                    \n");
                            htmlStr.Append("					<td colspan=5 class=xl1677651 width=291                                                                                                                                              \n");
                            htmlStr.Append("						style='border-right: .5pt solid black; border-left: none; width: 218pt'>&nbsp;" + dt_d.Rows[v_index][0] + "</td>                                                                                 \n");
                            htmlStr.Append("					<td class=xl1257651 style='border-left: none'>&nbsp;" + dt_d.Rows[v_index][1] + "</td>                                                                                                               \n");
                            htmlStr.Append("					<td colspan=2 class=xl1707651 style='border-left: none'>" + dt_d.Rows[v_index][2] + "</td>                                                                                                           \n");
                            htmlStr.Append("					<td class=xl1227651>&nbsp;</td>                                                                                                                                                      \n");
                            htmlStr.Append("					<td colspan=2 class=xl1407651 style='border-left: none'>" + dt_d.Rows[v_index][3] + "</td>                                                                                                           \n");
                            htmlStr.Append("					<td class=xl1217651>&nbsp;</td>                                                                                                                                                      \n");
                            htmlStr.Append("					<td colspan=2 class=xl1727651 width=133                                                                                                                                              \n");
                            htmlStr.Append("						style='border-left: none; width: 99pt'>" + dt_d.Rows[v_index][4] + "</td>                                                                                                                        \n");
                            htmlStr.Append("					<td class=xl1217651>&nbsp;</td>                                                                                                                                                      \n");
                            htmlStr.Append("				</tr>                                                                                                                                                                                    \n");

                        }
                        else
                        {
                            for (int i = 0; i < item_name.Count; i++)
                            {

                                if (!vlongItemName && item_name[i].ToString().Length >= 42)
                                {
                                    v_rowHeight = "26.0pt"; //"26.5pt";    
                                    v_rowHeightLast = "26.0pt"; //"27.5pt";
                                    v_rowHeightLastNumber = 22;//27.5;
                                    v_rowHeightEmptyLast = "22.0pt"; //"23.0pt";
                                    v_rowHeightNumber = 30;
                                    vlongItemName = true;
                                }
                                else
                                {
                                    v_rowHeight = "22.0pt";
                                    v_rowHeightNumber = 22;
                                }

                                if (k == v_countNumberOfPages - 1)
                                {
                                    v_totalHeightLastPage = v_totalHeightLastPage - v_rowHeightLastNumber;
                                }
                                else
                                {
                                    v_totalHeightPage = v_totalHeightPage - (v_rowHeightNumber);
                                }


                                if (i == 0)
                                {


                                    htmlStr.Append("							<tr class=xl757651 height=25                                                                                                                                                 \n");
                                    htmlStr.Append("								style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ";'>                                                                                                                    \n");
                                    htmlStr.Append("								<td colspan=2 rowspan=" + item_name.Count + "  height=25 class=xl1777651 width=39                                                                                          \n");
                                    htmlStr.Append("									style='border-right: .5pt solid black; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "; width: 29pt'>&nbsp;" + dt_d.Rows[v_index][7] + "</td>                                                                        \n");
                                    htmlStr.Append("								<td colspan=5 class=xl1677651 width=166 style='width: 125pt;border-bottom:none;border-top: 1.0pt dotted windowtext'>&nbsp;" + item_name[i] + "</td>                                         \n");
                                    htmlStr.Append("								<td class=xl1287651 rowspan=" + item_name.Count + " width=78 style=' width: 55.1pt'>&nbsp;" + dt_d.Rows[v_index][1] + "</td>                                                               \n");
                                    htmlStr.Append("								<td colspan=2 rowspan=" + item_name.Count + " class=xl1337651                                                                                                              \n");
                                    htmlStr.Append("									style='width: 21.85pt'>&nbsp;" + dt_d.Rows[v_index][2] + "</td>                                                                                                                      \n");
                                    htmlStr.Append("								<td class=xl1307651 rowspan=" + item_name.Count + " width=6 style=' width: 3.8pt'>&nbsp;</td>                                                                              \n");
                                    htmlStr.Append("								<td colspan=2  rowspan=" + item_name.Count + " class=xl1807651 style='border-left: none'>&nbsp;" + dt_d.Rows[v_index][3] + "</td>                                                          \n");
                                    htmlStr.Append("								<td class=xl1317651 rowspan=" + item_name.Count + " width=6 style=' width: 3.8pt'>&nbsp;</td>                                                                              \n");
                                    htmlStr.Append("								<td colspan=2 rowspan=" + item_name.Count + " class=xl1807651 style='border-left: none'>&nbsp;" + dt_d.Rows[v_index][4] + "</td>                                                           \n");
                                    htmlStr.Append("								<td class=xl1237651  rowspan=" + item_name.Count + "width=6 >&nbsp;</td>                                                                                                   \n");
                                    htmlStr.Append("							</tr>                                                                                                                                                                        \n");

                                }
                                else
                                {

                                    htmlStr.Append("							<tr class=xl757651 height=25                                                                                                                                                 \n");
                                    htmlStr.Append("								style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ";border-bottom: none'>                                                                                                 \n");
                                    htmlStr.Append("								<td colspan=5 class=xl1677651  width=166 style='width: 125pt;border-bottom: none;border-top: none;border-right: .5pt solid black;" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>&nbsp;" + item_name[i] + "</td>           \n");
                                    htmlStr.Append("								                                                                                                                                                                         \n");
                                    htmlStr.Append("		</tr>                                                                                                                                                                 \n");


                                }
                            }
                        }
                    }
                    else if (dtR == page[k] - 1)
                    {

                        if (item_name.Count == 1)
                        {
                            if (k == v_countNumberOfPages - 1)
                            {
                                v_totalHeightLastPage = v_totalHeightLastPage - v_rowHeightLastNumber;
                            }
                            else
                            {
                                v_totalHeightPage = v_totalHeightPage - (v_rowHeightNumber);
                            }

                            htmlStr.Append("				<tr class=xl757651 height=25                                                                                                                                                             \n");
                            htmlStr.Append("					style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ";border-bottom: .5pt solid black;'>                                                                                                \n");
                            htmlStr.Append("					<td colspan=2 height=25 class=xl1777651 width=39                                                                                                                                     \n");
                            htmlStr.Append("						style='border-right: .5pt solid black; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "; width: 29pt'>&nbsp;" + dt_d.Rows[v_index][7] + "</td>                                                                                    \n");
                            htmlStr.Append("					<td colspan=5 class=xl1797651 width=166 style='width: 125pt'>&nbsp;" + dt_d.Rows[v_index][0] + "</td>                                                                                                \n");
                            htmlStr.Append("					<td class=xl1287651 width=78 style=' width: 55.1pt'>&nbsp;" + dt_d.Rows[v_index][1] + "</td>                                                                                                         \n");
                            htmlStr.Append("					<td colspan=2 class=xl1337651                                                                                                                                                        \n");
                            htmlStr.Append("						style='width: 21.85pt'>&nbsp;" + dt_d.Rows[v_index][2] + "</td>                                                                                                                                  \n");
                            htmlStr.Append("					<td class=xl1307651 width=6 style=' width: 3.8pt'>&nbsp;</td>                                                                                                                        \n");
                            htmlStr.Append("					<td colspan=2 class=xl1807651 style='border-left: none'>&nbsp;" + dt_d.Rows[v_index][3] + "</td>                                                                                                     \n");
                            htmlStr.Append("					<td class=xl1317651 width=6 style=' width: 3.8pt'>&nbsp;</td>                                                                                                                        \n");
                            htmlStr.Append("					<td colspan=2 class=xl1807651 style='border-left: none'>&nbsp;" + dt_d.Rows[v_index][4] + "</td>                                                                                                     \n");
                            htmlStr.Append("					<td class=xl1237651 width=6 >&nbsp;</td>                                                                                                                                             \n");
                            htmlStr.Append("				</tr>                                                                                                                                                                                    \n");

                        }
                        else
                        {
                            for (int i = 0; i < item_name.Count; i++)
                            {
                                if (!vlongItemName && item_name[i].ToString().Length >= 42)
                                {
                                    v_rowHeight = "26.0pt"; //"26.5pt";    
                                    v_rowHeightLast = "26.0pt"; //"27.5pt";
                                    v_rowHeightLastNumber = 22;//27.5;
                                    v_rowHeightEmptyLast = "22.0pt"; //"23.0pt";
                                    v_rowHeightNumber = 30;
                                    vlongItemName = true;
                                }
                                else if (i == (item_name.Count - 1) && item_name[i].ToString().Length >= 42)
                                {
                                    v_rowHeightNumber = 35;
                                    v_rowHeight = "28.0pt"; //"26.5pt";    
                                    v_rowHeightLast = "28.0pt"; //"27.5pt";
                                }
                                else
                                {
                                    v_rowHeight = "22.0pt";
                                    v_rowHeightNumber = 22;
                                }
                                if (k == v_countNumberOfPages - 1)
                                {
                                    v_totalHeightLastPage = v_totalHeightLastPage - v_rowHeightLastNumber;
                                }
                                else
                                {
                                    v_totalHeightPage = v_totalHeightPage - (v_rowHeightNumber);
                                }

                                if (i == 0)
                                {

                                    htmlStr.Append("							<tr class=xl757651 height=25                                                                                                                                                 \n");
                                    htmlStr.Append("								style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ";'>                                                                                                                    \n");
                                    htmlStr.Append("								<td colspan=2 rowspan=" + item_name.Count + "  height=25 class=xl1777651 width=39                                                                                          \n");
                                    htmlStr.Append("									style='border-right: .5pt solid black; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "; width: 29pt'>&nbsp;" + dt_d.Rows[v_index][7] + "</td>                                                                        \n");
                                    htmlStr.Append("								<td colspan=5 class=xl1677651 width=166 style='width: 125pt;border-bottom:none;border-top: 1.0pt dotted windowtext'>&nbsp;" + item_name[i] + "</td>                                         \n");
                                    htmlStr.Append("								<td class=xl1287651 rowspan=" + item_name.Count + " width=78 style=' width: 55.1pt'>&nbsp;" + dt_d.Rows[v_index][1] + "</td>                                                               \n");
                                    htmlStr.Append("								<td colspan=2 rowspan=" + item_name.Count + " class=xl1337651                                                                                                              \n");
                                    htmlStr.Append("									style='width: 21.85pt'>&nbsp;" + dt_d.Rows[v_index][2] + "</td>                                                                                                                      \n");
                                    htmlStr.Append("								<td class=xl1307651 rowspan=" + item_name.Count + " width=6 style=' width: 3.8pt'>&nbsp;</td>                                                                              \n");
                                    htmlStr.Append("								<td colspan=2  rowspan=" + item_name.Count + " class=xl1807651 style='border-left: none'>&nbsp;" + dt_d.Rows[v_index][3] + "</td>                                                          \n");
                                    htmlStr.Append("								<td class=xl1317651 rowspan=" + item_name.Count + " width=6 style=' width: 3.8pt'>&nbsp;</td>                                                                              \n");
                                    htmlStr.Append("								<td colspan=2 rowspan=" + item_name.Count + " class=xl1807651 style='border-left: none'>&nbsp;" + dt_d.Rows[v_index][4] + "</td>                                                           \n");
                                    htmlStr.Append("								<td class=xl1237651  rowspan=" + item_name.Count + "width=6 >&nbsp;</td>                                                                                                   \n");
                                    htmlStr.Append("							</tr>	                                                                                                                                                                     \n");

                                }
                                else
                                {


                                    htmlStr.Append("							<tr class=xl757651 height=25                                                                                                                                                 \n");
                                    htmlStr.Append("								style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ";border-bottom: none'>                                                                                                 \n");
                                    htmlStr.Append("								<td colspan=5 class=xl1677651 width=166 style='width: 125pt;border-bottom: none;border-top: none;border-right: .5pt solid black;" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>&nbsp;" + item_name[i] + "</td>            \n");
                                    htmlStr.Append("								                                                                                                                                                                         \n");
                                    htmlStr.Append("							</tr>                                                                                                                                                                        \n");


                                }
                            }
                        }
                    }
                    else
                    {

                        if (item_name.Count == 1)
                        {
                            if (k == v_countNumberOfPages - 1)
                            {
                                v_totalHeightLastPage = v_totalHeightLastPage - v_rowHeightLastNumber;
                            }
                            else
                            {
                                v_totalHeightPage = v_totalHeightPage - (v_rowHeightNumber);
                            }

                            htmlStr.Append("					<tr class=xl837651 height=25                                                                                                                                                         \n");
                            htmlStr.Append("						style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                                                                             \n");
                            htmlStr.Append("						<td colspan=2 height=25 class=xl1657651 width=39                                                                                                                                 \n");
                            htmlStr.Append("							style='border-right: .5pt solid black; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "; width: 29pt'>&nbsp;" + dt_d.Rows[v_index][7] + "</td>                                                                                \n");
                            htmlStr.Append("						<td colspan=5 class=xl1677651 width=291                                                                                                                                          \n");
                            htmlStr.Append("							style='border-right: .5pt solid black; border-left: none; width: 218pt'>&nbsp;" + dt_d.Rows[v_index][0] + "</td>                                                                             \n");
                            htmlStr.Append("						<td class=xl1257651 style=' border-left: none'>&nbsp;" + dt_d.Rows[v_index][1] + "</td>                                                                                                          \n");
                            htmlStr.Append("						<td colspan=2 class=xl1417651 >&nbsp;" + dt_d.Rows[v_index][2] + "</td>                                                                                                                          \n");
                            htmlStr.Append("						<td class=xl1227651 >&nbsp;</td>                                                                                                                                                 \n");
                            htmlStr.Append("						<td colspan=2 class=xl1407651                                                                                                                                                    \n");
                            htmlStr.Append("							style=' border-left: none'>&nbsp;" + dt_d.Rows[v_index][3] + "</td>                                                                                                                          \n");
                            htmlStr.Append("						<td class=xl1217651 >&nbsp;</td>                                                                                                                                                 \n");
                            htmlStr.Append("						<td colspan=2 class=xl1727651 width=133                                                                                                                                          \n");
                            htmlStr.Append("							style='border-left: none; width: 99pt'>&nbsp;" + dt_d.Rows[v_index][4] + "</td>                                                                                                              \n");
                            htmlStr.Append("						<td class=xl1217651 style='border-top: none'>&nbsp;</td>                                                                                                                         \n");
                            htmlStr.Append("					</tr>                                                                                                                                                                                \n");

                        }
                        else
                        {
                            for (int i = 0; i < item_name.Count; i++)
                            {
                                if (!vlongItemName && item_name[i].ToString().Length >= 42)
                                {
                                    v_rowHeight = "26.0pt"; //"26.5pt";    
                                    v_rowHeightLast = "26.0pt"; //"27.5pt";
                                    v_rowHeightLastNumber = 22;//27.5;
                                    v_rowHeightEmptyLast = "22.0pt"; //"23.0pt";
                                    v_rowHeightNumber = 30;
                                    vlongItemName = true;
                                }
                                else if (i == (item_name.Count - 1) && item_name[i].ToString().Length >= 42)
                                {
                                    v_rowHeightNumber = 29;
                                    v_rowHeight = "28.0pt"; //"26.5pt";    
                                    v_rowHeightLast = "28.0pt"; //"27.5pt";
                                }
                                else
                                {
                                    v_rowHeight = "22.0pt";
                                    v_rowHeightNumber = 22;
                                }
                                if (k == v_countNumberOfPages - 1)
                                {
                                    v_totalHeightLastPage = v_totalHeightLastPage - v_rowHeightLastNumber;
                                }
                                else
                                {
                                    v_totalHeightPage = v_totalHeightPage - (v_rowHeightNumber);
                                }

                                if (i == 0)
                                {

                                    htmlStr.Append("							<tr class=xl757651 height=25                                                                                                                                                 \n");
                                    htmlStr.Append("								style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ";'>                                                                                                                    \n");
                                    htmlStr.Append("								<td colspan=2 rowspan=" + item_name.Count + "  height=25 class=xl1777651 width=39                                                                                          \n");
                                    htmlStr.Append("									style='border-right: .5pt solid black; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "; width: 29pt'>&nbsp;" + dt_d.Rows[v_index][7] + "</td>                                                                        \n");
                                    htmlStr.Append("								<td colspan=5 class=xl1677651  width=166 style='width: 125pt;border-bottom:none;border-top:1.0pt dotted windowtext'>&nbsp;" + item_name[i] + "</td>                                         \n");
                                    htmlStr.Append("								<td class=xl1287651 rowspan=" + item_name.Count + " width=78 style=' width: 55.1pt'>&nbsp;" + dt_d.Rows[v_index][1] + "</td>                                                               \n");
                                    htmlStr.Append("								<td colspan=2 rowspan=" + item_name.Count + " class=xl1337651                                                                                                              \n");
                                    htmlStr.Append("									style='width: 21.85pt'>&nbsp;" + dt_d.Rows[v_index][2] + "</td>                                                                                                                      \n");
                                    htmlStr.Append("								<td class=xl1307651 rowspan=" + item_name.Count + " width=6 style=' width: 3.8pt'>&nbsp;</td>                                                                              \n");
                                    htmlStr.Append("								<td colspan=2  rowspan=" + item_name.Count + " class=xl1807651 style='border-left: none'>&nbsp;" + dt_d.Rows[v_index][3] + "</td>                                                          \n");
                                    htmlStr.Append("								<td class=xl1317651 rowspan=" + item_name.Count + " width=6 style=' width: 3.8pt'>&nbsp;</td>                                                                              \n");
                                    htmlStr.Append("								<td colspan=2 rowspan=" + item_name.Count + " class=xl1807651 style='border-left: none'>&nbsp;" + dt_d.Rows[v_index][4] + "</td>                                                           \n");
                                    htmlStr.Append("								<td class=xl1237651  rowspan=" + item_name.Count + "width=6 >&nbsp;</td>                                                                                                   \n");
                                    htmlStr.Append("							</tr>                                                                                                                                                                        \n");

                                }
                                else
                                {

                                    htmlStr.Append("							<tr class=xl757651 height=25                                                                                                                                                 \n");
                                    htmlStr.Append("								style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ";border-bottom: none'>                                                                                                 \n");
                                    htmlStr.Append("								<td colspan=5 class=xl1677651  width=166 style='width: 125pt;border-bottom: none;border-top: none;border-right: .5pt solid black;" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>&nbsp;" + item_name[i] + "</td>           \n");
                                    htmlStr.Append("								                                                                                                                                                                         \n");
                                    htmlStr.Append("							</tr>                                                                                                                                                                        \n");

                                }
                            }
                        }
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
                    v_spacePerPage = v_totalHeightPage;//  18;
                }
                else if (k < v_countNumberOfPages - 1 && page_index[k] > rowsPerPage)
                {
                    v_spacePerPage = v_totalHeightPage;//  18;
                }
                if (k == v_countNumberOfPages - 1 && page_index[k] < rowsPerPage) // Trang cuoi khong du dong
                {
                    v_rowHeightEmptyLast = Math.Round(v_totalHeightLastPage / (rowsPerPage - page_index[k]), 2).ToString() + "pt";
                    for (int i = 0; i < rowsPerPage - page_index[k]; i++)
                    {
                        if (i == (rowsPerPage - page_index[k] - 1))
                        {

                            htmlStr.Append("							<tr class=xl757651 height=25                                                                                                                                                 \n");
                            htmlStr.Append("					style='mso-height-source: userset; height: " + v_rowHeightEmptyLast + "'>                                                                                                                                 \n");
                            htmlStr.Append("					<td colspan=2 height=25 class=xl1777651 width=39                                                                                                                                     \n");
                            htmlStr.Append("						style='border-right: .5pt solid black; height: " + v_rowHeightEmptyLast + "; width: 29pt'>&nbsp;</td>                                                                                                 \n");
                            htmlStr.Append("					<td colspan=5 class=xl1797651 width=166 style='width: 125pt'>&nbsp;</td>                                                                                                             \n");
                            htmlStr.Append("					<td class=xl1287651 width=78 style=' width: 55.1pt'>&nbsp;</td>                                                                                                                      \n");
                            htmlStr.Append("					<td colspan=2 class=xl1337651                                                                                                                                                        \n");
                            htmlStr.Append("						style=' width: 21.85pt'>&nbsp;</td>                                                                                                                                              \n");
                            htmlStr.Append("					<td class=xl1307651 width=6 style='width: 3.8pt'>&nbsp;</td>                                                                                                                         \n");
                            htmlStr.Append("					<td colspan=2 class=xl1807651 style='border-left: none'>&nbsp;</td>                                                                                                                  \n");
                            htmlStr.Append("					<td class=xl1317651 width=6 style='width: 3.8pt'>&nbsp;</td>                                                                                                                         \n");
                            htmlStr.Append("					<td colspan=2 class=xl1807651 style='border-left: none'>&nbsp;</td>                                                                                                                  \n");
                            htmlStr.Append("					<td class=xl1237651 width=6 >&nbsp;</td>                                                                                                                                             \n");
                            htmlStr.Append("				</tr>                                                                                                                                                                                    \n");

                        }
                        else
                        {

                            htmlStr.Append("						<tr class=xl837651 height=25                                                                                                                                                     \n");
                            htmlStr.Append("					style='mso-height-source: userset; height: " + v_rowHeightEmptyLast + "'>                                                                                                                                 \n");
                            htmlStr.Append("					<td colspan=2 height=25 class=xl1657651 width=39                                                                                                                                     \n");
                            htmlStr.Append("						style='border-right: .5pt solid black; height: " + v_rowHeightEmptyLast + "; width: 29pt'>&nbsp;</td>                                                                                                 \n");
                            htmlStr.Append("					<td colspan=5 class=xl1677651 width=291                                                                                                                                              \n");
                            htmlStr.Append("						style='border-right: .5pt solid black; border-left: none; width: 218pt'>&nbsp;</td>                                                                                              \n");
                            htmlStr.Append("					<td class=xl1257651 style=' border-left: none'>&nbsp;</td>                                                                                                                           \n");
                            htmlStr.Append("					<td colspan=2 class=xl1417651 >&nbsp;</td>                                                                                                                                           \n");
                            htmlStr.Append("					<td class=xl1227651 >&nbsp;</td>                                                                                                                                                     \n");
                            htmlStr.Append("					<td colspan=2 class=xl1407651                                                                                                                                                        \n");
                            htmlStr.Append("						style=' border-left: none'>&nbsp;</td>                                                                                                                                           \n");
                            htmlStr.Append("					<td class=xl1217651 >&nbsp;</td>                                                                                                                                                     \n");
                            htmlStr.Append("					<td colspan=2 class=xl1727651 width=133                                                                                                                                              \n");
                            htmlStr.Append("						style='border-left: none; width: 99pt'>&nbsp;</td>                                                                                                                               \n");
                            htmlStr.Append("					<td class=xl1217651 ></td>                                                                                                                                                           \n");
                            htmlStr.Append("				</tr>                                                                                                                                                                                    \n");


                        }
                    } // for de du dong                                                                                                                                                 
                }  // if k == v_countNumberOfPages - 1 && vPage[k] < pos			

                if (k < v_countNumberOfPages - 1) // phan cuoi cung cua trang( pos_vl = 20) 
                {

                    htmlStr.Append("       <tr class=xl767652 height = 25   \n");
                    htmlStr.Append("                    style='mso-height-source: userset; height: 18pt;border-bottom: .5pt solid black;' >                                                               \n");
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
                    htmlStr.Append("	</table>             																																										\n");
                    htmlStr.Append("	<table  border=0>                                                                                                                                                                                                 \n");
                    htmlStr.Append("		<tr height=5  style='height: " + (v_spacePerPage).ToString() + "pt'>                                                                                                                                                                \n");
                    htmlStr.Append("		</tr>      																																														\n");
                    htmlStr.Append("	</table>             																																										\n");

                }

            }

            htmlStr.Append("		<tr class=xl757651 height=26                                                                                                                                                                    \n");
            htmlStr.Append("			style='mso-height-source: userset; height: 22.0pt'>                                                                                                                                         \n");
            htmlStr.Append("			<td height=26 class=xl1037651 width=6 colspan=7                                                                                                                                                       \n");
            htmlStr.Append("				style='height: 22.0pt; width: 4.75pt'>&nbsp;" + lb_amount_trans + "</td>                                                                                                                                        \n");
            //htmlStr.Append("			<td class=xl1327651 width=33 style='width: 29.68pt'>&nbsp;</td>                                                                                                                             \n");
            //htmlStr.Append("			<td class=xl1327651 width=70 style='width: 49.4pt'>&nbsp;</td>                                                                                                                              \n");
            //htmlStr.Append("			<td class=xl1327651 width=55 style='width: 48.68pt'>&nbsp;</td>                                                                                                                             \n");
            //h/tmlStr.Append("			<td class=xl1327651 width=41 style='width: 36.81pt'>&nbsp;</td>                                                                                                                             \n");
            //htmlStr.Append("			<td class=xl1327651 width=98 style='width: 87.88pt'>&nbsp;</td>                                                                                                                              \n");
            //htmlStr.Append("			<td class=xl1327651 width=27 style='width: 23.75pt'>&nbsp;</td>                                                                                                                                \n");
            htmlStr.Append("			<td colspan=7 class=xl1847651 width=273                                                                                                                                                     \n");
            htmlStr.Append("				style='border-right: 0.5pt solid black; width: 204pt'>C&#7897;ng                                                                                                                         \n");
            htmlStr.Append("				ti&#7873;n hàng (<font class='font107651'>Total amount</font><font                                                                                                                      \n");
            htmlStr.Append("				class='font87651'>) :</font>                                                                                                                                                            \n");
            htmlStr.Append("			</td>                                                                                                                                                                                       \n");
            htmlStr.Append("			<td colspan=2 class=xl1867651 style='border-left: none'>&nbsp;" + amount_net + "</td>                                                                                                           \n");
            htmlStr.Append("			<td class=xl1247651>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
            htmlStr.Append("		<tr class=xl757651 height=26                                                                                                                                                                    \n");
            htmlStr.Append("			style='mso-height-source: userset; height: 24.0pt'>                                                                                                                                         \n");
            htmlStr.Append("			<td height=26 class=xl1037651 width=6                                                                                                                                                       \n");
            htmlStr.Append("				style='height: 24.0pt; border-top: none; width: 4.75pt'>&nbsp;</td>                                                                                                                      \n");
            htmlStr.Append("			<td colspan=4 class=xl1427651 width=199 style='width: 149pt'><span                                                                                                                          \n");
            htmlStr.Append("				style='mso-spacerun: yes'> </span>Thu&#7871; su&#7845;t GTGT (<font                                                                                                                     \n");
            htmlStr.Append("				class='font107651'>VAT rate</font><font class='font87651'>)                                                                                                                             \n");
            htmlStr.Append("					:</font></td>                                                                                                                                                                       \n");
            htmlStr.Append("			<td class=xl1357651 width=98 style='border-top: none; width: 87.88pt'>&nbsp;" + dt.Rows[0]["taxrate"] + "</td>                                                                                             \n");
            htmlStr.Append("			<td class=xl1427651 width=27 style='border-top: none; width: 23.75pt'>&nbsp;</td>                                                                                                              \n");
            htmlStr.Append("			<td colspan=7 class=xl1847651 width=273                                                                                                                                                     \n");
            htmlStr.Append("				style='border-right: 0.5pt solid black; width: 204pt'>Ti&#7873;n                                                                                                                         \n");
            htmlStr.Append("				thu&#7871; GTGT (<font class='font107651'>VAT</font><font                                                                                                                               \n");
            htmlStr.Append("				class='font87651'>) :</font>                                                                                                                                                            \n");
            htmlStr.Append("			</td>                                                                                                                                                                                       \n");
            htmlStr.Append("			<td colspan=2 class=xl1867651 style='border-left: none'>&nbsp;" + amount_vat + "</td>                                                                                                         \n");
            htmlStr.Append("			<td class=xl877651 style='border-top: none'>&nbsp;</td>                                                                                                                                     \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
            htmlStr.Append("		<tr class=xl757651 height=26                                                                                                                                                                    \n");
            htmlStr.Append("			style='mso-height-source: userset; height: 24.0pt'>                                                                                                                                         \n");
            htmlStr.Append("			<td height=26 class=xl1037651 width=6                                                                                                                                                       \n");
            htmlStr.Append("				style='height: 24.0pt; border-top: none; width: 4.75pt'>&nbsp;</td>                                                                                                                      \n");
            htmlStr.Append("	 <td colspan=3 class=xl1427651 width=199 style='width: 149pt'><span															  \n");
            htmlStr.Append("	 					style='mso-spacerun: yes'> </span>T&#7893;ng ti&#7873;n (<font                                            \n");
            htmlStr.Append("	 					class='font107651'>Amount VND</font><font class='font87651'>)                                                 \n");
            htmlStr.Append("	 						:</font></td>                                                                                         \n");
            htmlStr.Append("	 <td colspan=2 class=xl1427651 width=98 style='border-top: none; width: 70.3pt'>&nbsp;" + amount_trans + "</td>                \n");
            htmlStr.Append("			<td class=xl1427651 width=27 style='border-top: none; width: 23.75pt'>&nbsp;</td>                                                                                                              \n");
            htmlStr.Append("			<td colspan=7 class=xl1847651 width=273                                                                                                                                                     \n");
            htmlStr.Append("				style='border-right: 0.5pt solid black; width: 204pt'>T&#7893;ng                                                                                                                         \n");
            htmlStr.Append("				ti&#7873;n thanh toán (<font class='font107651'>Total                                                                                                                                   \n");
            htmlStr.Append("					payment</font><font class='font87651'>) :</font>                                                                                                                                    \n");
            htmlStr.Append("			</td>                                                                                                                                                                                       \n");
            htmlStr.Append("			<td colspan=2 class=xl1867651 style='border-left: none'>&nbsp;" + amount_total + "</td>                                                                                                  \n");
            htmlStr.Append("			<td class=xl877651>&nbsp;</td>                                                                                                                                                              \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
            htmlStr.Append("		<tr class=xl757651 height=24                                                                                                                                                                    \n");
            htmlStr.Append("			style='mso-height-source: userset; height: 20.0pt'>                                                                                                                                         \n");
            htmlStr.Append("			<td height=24 class=xl867651                                                                                                                                                                \n");
            htmlStr.Append("				style='height: 20.0pt; border-top: none'>&nbsp;</td>                                                                                                                                    \n");
            htmlStr.Append("			<td colspan=15 class=xl1977651>S&#7889; ti&#7873;n vi&#7871;t                                                                                                                               \n");
            htmlStr.Append("				b&#7857;ng ch&#7919; (<font class='font107651'>In words</font><font                                                                                                                     \n");
            htmlStr.Append("				class='font67651'>): " + read_prive + "</font>                                                                                                                                             \n");
            htmlStr.Append("			</td>                                                                                                                                                                                       \n");
            htmlStr.Append("			<td class=xl887651 width=6 style='width: 4.75pt'>&nbsp;</td>                                                                                                                                 \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
            htmlStr.Append("		<tr height=18 style='mso-height-source: userset; height: 6.95pt'>                                                                                                                              \n");
            htmlStr.Append("			<td height=18 class=xl827651 style='height: 6.95pt'>&nbsp;</td>                                                                                                                            \n");
            htmlStr.Append("			<td class=xl897651>&nbsp;</td>                                                                                                                                                              \n");
            htmlStr.Append("			<td class=xl897651>&nbsp;</td>                                                                                                                                                              \n");
            htmlStr.Append("			<td class=xl897651>&nbsp;</td>                                                                                                                                                              \n");
            htmlStr.Append("			<td class=xl907651 width=41 style='width: 36.81pt'>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("			<td class=xl907651 width=98 style='width: 87.88pt'>&nbsp;</td>                                                                                                                               \n");
            htmlStr.Append("			<td class=xl907651 width=27 style='width: 23.75pt'>&nbsp;</td>                                                                                                                                 \n");
            htmlStr.Append("			<td class=xl907651 width=78 style='width: 68.88pt'>&nbsp;</td>                                                                                                                               \n");
            htmlStr.Append("			<td class=xl907651 width=56 style='width: 49.88pt'>&nbsp;</td>                                                                                                                               \n");
            htmlStr.Append("			<td class=xl907651 width=30 style='width: 27.32pt'>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("			<td class=xl907651 width=6 style='width: 4.75pt'>&nbsp;</td>                                                                                                                                 \n");
            htmlStr.Append("			<td class=xl907651 width=49 style='width: 43.94pt'>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("			<td class=xl907651 width=48 style='width: 42.75pt'>&nbsp;</td>                                                                                                                               \n");
            htmlStr.Append("			<td class=xl907651 width=6 style='width: 4.75pt'>&nbsp;</td>                                                                                                                                 \n");
            htmlStr.Append("			<td class=xl907651 width=55 style='width: 48.68pt'>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("			<td class=xl907651 width=78 style='width: 68.88pt'>&nbsp;</td>                                                                                                                               \n");
            htmlStr.Append("			<td class=xl917651 width=6 style='width: 4.75pt'>&nbsp;</td>                                                                                                                                 \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
            htmlStr.Append("		<tr class=xl797651 height=24                                                                                                                                                                    \n");
            htmlStr.Append("			style='mso-height-source: userset; height: 24.0pt'>                                                                                                                                         \n");
            htmlStr.Append("			<td height=24 class=xl947651 style='height: 24.0pt'>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("			<td colspan=5 class=xl1987651 width=297 style='width: 223pt'>Ng&#432;&#7901;i                                                                                                               \n");
            htmlStr.Append("				mua hàng (Buyer)</td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl927651 width=27 style='width: 23.75pt'>&nbsp;</td>                                                                                                                                 \n");
            htmlStr.Append("			<td class=xl927651 width=78 style='width: 68.88pt'>&nbsp;</td>                                                                                                                               \n");
            htmlStr.Append("			<td colspan=8 class=xl1987651 width=328 style='width: 245pt'>Ng&#432;&#7901;i                                                                                                               \n");
            htmlStr.Append("				bán hàng (Seller)</td>                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl937651 width=6 style='width: 4.75pt'>&nbsp;</td>                                                                                                                                 \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
            htmlStr.Append("		<tr class=xl697651 height=18                                                                                                                                                                    \n");
            htmlStr.Append("			style='mso-height-source: userset; height: 13.95pt'>                                                                                                                                        \n");
            htmlStr.Append("			<td height=18 class=xl957651 style='height: 13.95pt'>&nbsp;</td>                                                                                                                            \n");
            htmlStr.Append("			<td colspan=5 class=xl1827651 width=297 style='width: 223pt'>(Ký,                                                                                                                           \n");
            htmlStr.Append("				ghi rõ h&#7885; tên)</td>                                                                                                                                                               \n");
            htmlStr.Append("			<td class=xl1137651 width=27 style='width: 23.75pt'>&nbsp;</td>                                                                                                                                \n");
            htmlStr.Append("			<td class=xl1137651 width=78 style='width: 68.88pt'>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("			<td colspan=8 class=xl1827651 width=328 style='width: 245pt'>(Ký,                                                                                                                           \n");
            htmlStr.Append("				&#273;óng d&#7845;u, ghi rõ h&#7885; tên)</td>                                                                                                                                          \n");
            htmlStr.Append("			<td class=xl967651 width=6 style='width: 4.75pt'>&nbsp;</td>                                                                                                                                 \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
            htmlStr.Append("		<tr class=xl997651 height=20                                                                                                                                                                    \n");
            htmlStr.Append("			style='mso-height-source: userset; height: 15.45pt'>                                                                                                                                        \n");
            htmlStr.Append("			<td height=20 class=xl977651 style='height: 15.45pt'>&nbsp;</td>                                                                                                                            \n");
            htmlStr.Append("			<td colspan=5 class=xl1837651 width=297 style='width: 223pt'>(Signature                                                                                                                     \n");
            htmlStr.Append("				&amp; full name)</td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl1127651 width=27 style='width: 23.75pt'>&nbsp;</td>                                                                                                                                \n");
            htmlStr.Append("			<td class=xl1127651 width=78 style='width: 68.88pt'>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("			<td colspan=8 class=xl1837651 width=328 style='width: 245pt'>(Signature,                                                                                                                    \n");
            htmlStr.Append("				stamp &amp; full name)</td>                                                                                                                                                             \n");
            htmlStr.Append("			<td class=xl987651 width=6 style='width: 4.75pt'>&nbsp;</td>                                                                                                                                 \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
            htmlStr.Append("		<tr height=21 style='height: 6.5pt'>                                                                                                                                                           \n");
            htmlStr.Append("			<td height=21 class=xl1377651 style='height: 6.5pt'>&nbsp;</td>                                                                                                                            \n");
            htmlStr.Append("			<td class=xl657651>&nbsp;</td>                                                                                                                                                              \n");
            htmlStr.Append("			<td class=xl657651>&nbsp;</td>                                                                                                                                                              \n");
            htmlStr.Append("			<td class=xl657651>&nbsp;</td>                                                                                                                                                              \n");
            htmlStr.Append("			<td class=xl717651 width=41 style='width: 36.81pt'>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("			<td class=xl717651 width=98 style='width: 87.88pt'>&nbsp;</td>                                                                                                                               \n");
            htmlStr.Append("			<td class=xl717651 width=27 style='width: 23.75pt'>&nbsp;</td>                                                                                                                                 \n");
            htmlStr.Append("			<td class=xl717651 width=78 style='width: 68.88pt'>&nbsp;</td>                                                                                                                               \n");
            htmlStr.Append("			<td class=xl717651 width=56 style='width: 49.88pt'>&nbsp;</td>                                                                                                                               \n");
            htmlStr.Append("			<td class=xl717651 width=30 style='width: 27.32pt'>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("			<td class=xl717651 width=6 style='width: 4.75pt'>&nbsp;</td>                                                                                                                                 \n");
            htmlStr.Append("			<td class=xl717651 width=49 style='width: 43.94pt'>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("			<td class=xl717651 width=48 style='width: 42.75pt'>&nbsp;</td>                                                                                                                               \n");
            htmlStr.Append("			<td class=xl717651 width=6 style='width: 4.75pt'>&nbsp;</td>                                                                                                                                 \n");
            htmlStr.Append("			<td class=xl717651 width=55 style='width: 48.68pt'>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("			<td class=xl717651 width=78 style='width: 68.88pt'>&nbsp;</td>                                                                                                                               \n");
            htmlStr.Append("			<td class=xl1007651 width=6 style='width: 4.75pt'>&nbsp;</td>                                                                                                                                \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
            htmlStr.Append("		<tr height=22 style='height: 12.8pt'>                                                                                                                                                           \n");
            htmlStr.Append("			<td height=22 class=xl1377651 style='height: 12.8pt'>&nbsp;</td>                                                                                                                            \n");
            htmlStr.Append("			<td class=xl657651>&nbsp;</td>                                                                                                                                                              \n");
            htmlStr.Append("			<td class=xl657651>&nbsp;</td>                                                                                                                                                              \n");
            htmlStr.Append("			<td class=xl657651>&nbsp;</td>                                                                                                                                                              \n");
            htmlStr.Append("			<td class=xl717651 width=41 style='width: 36.81pt'>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("			<td class=xl717651 width=98 style='width: 87.88pt'>&nbsp;</td>                                                                                                                               \n");
            htmlStr.Append("			<td class=xl717651 width=27 style='width: 23.75pt'>&nbsp;</td>                                                                                                                                 \n");
            htmlStr.Append("			<td class=xl717651 width=78 style='width: 68.88pt'>&nbsp;</td>                                                                                                                               \n");
            htmlStr.Append("			<td class=xl1477651 colspan=4>&nbsp;Signature Valid</td>                                                                                                                                    \n");

            if (dt.Rows[0]["sign_yn"].ToString() == "Y")
            {

                htmlStr.Append("			<td align=left class=xl1067651 valign=top><![if !vml]><span                                                                                                                                                 \n");
                htmlStr.Append("				style='mso-ignore: vglayout; position: absolute; z-index: 1; margin-left: 18px; margin-top: 7px; width: 81px; height: 54px'><img                                                        \n");
                htmlStr.Append("					width=81 height=54                                                                                                                                                                  \n");
                htmlStr.Append("					src='D:\\webproject\\e-invoice-ws\\02.Web\\EInvoice\\img\\check_signed.png'                                                                                                                \n");
                htmlStr.Append("					v:shapes='Picture_x0020_8'></span>                                                                                                                                                  \n");
                htmlStr.Append("			<![endif]><span style='mso-ignore: vglayout2'>                                                                                                                                              \n");
                htmlStr.Append("					<table cellpadding=0 cellspacing=0>                                                                                                                                                 \n");
                htmlStr.Append("						<tr>                                                                                                                                                                            \n");
                htmlStr.Append("							<td height=22 class=xl1067651 width=48                                                                                                                                      \n");
                htmlStr.Append("								style='height: 16.8pt; width: 42.75pt'>&nbsp;</td>                                                                                                                       \n");
                htmlStr.Append("						</tr>                                                                                                                                                                           \n");
                htmlStr.Append("					</table>                                                                                                                                                                            \n");
                htmlStr.Append("			</span></td>                                                                                                                                                                                \n");
                htmlStr.Append("                                                                                                                                                                                                        \n");

            }
            else
            {

                htmlStr.Append("			<td height=22 class=xl1067651 width=48                                                                                                                                                      \n");
                htmlStr.Append("				style='height: 16.8pt; width: 42.75pt'>&nbsp;</td>                                                                                                                                       \n");

            }

            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append("			<td class=xl1077651>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("			<td class=xl1077651>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("			<td class=xl1087651>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("			<td class=xl1107651>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
            htmlStr.Append("		<tr class=xl1167651 height=32                                                                                                                                                                   \n");
            htmlStr.Append("			style='mso-height-source: userset; height: 24.45pt'>                                                                                                                                        \n");
            htmlStr.Append("			<td height=32 class=xl1157651 style='height: 24.45pt'>&nbsp;</td>                                                                                                                           \n");
            htmlStr.Append("			<td class=xl1167651>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("			<td class=xl1167651>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("			<td class=xl1167651>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("			<td class=xl1177651>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("			<td class=xl1177651>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("			<td class=xl1177651>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("			<td class=xl1177651>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("                                                                                                                                                                                                        \n");
            htmlStr.Append("			<td colspan=2 class=xl1887651><font class='font187651'>&#272;&#432;&#7907;c                                                                                                                 \n");
            htmlStr.Append("					ký b&#7903;i:</font></td>                                                                                                                                                           \n");
            htmlStr.Append("			<td class=xl1447651>&nbsp;</td>                                                                                                                                                             \n");

            if (dt.Rows[0]["sign_yn"].ToString() == "Y")
            {

                htmlStr.Append("			<td colspan=5 class=xl1907651 width=236                                                                                                                                                     \n");
                htmlStr.Append("				style='border-right: 0.5pt solid black; width: 176pt;font-size: 12.06pt'>" + dt.Rows[0]["SignedBy"] + "</td>                                                                                                            \n");


            }
            else
            {

                htmlStr.Append("			<td colspan=5 class=xl1907651 width=236                                                                                                                                                     \n");
                htmlStr.Append("				style='border-right: 0.5pt solid black; width: 176pt'></td>                                                                                                                              \n");

            }

            htmlStr.Append("			<td class=xl1187651>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
            htmlStr.Append("		<tr height=21 style='height: 19.5pt'>                                                                                                                                                           \n");
            htmlStr.Append("			<td height=21 class=xl1377651 style='height: 19.5pt'>&nbsp;</td>                                                                                                                            \n");
            htmlStr.Append("			<td class=xl657651>&nbsp;</td>                                                                                                                                                              \n");
            htmlStr.Append("			<td class=xl657651>&nbsp;</td>                                                                                                                                                              \n");
            htmlStr.Append("			<td class=xl657651>&nbsp;</td>                                                                                                                                                              \n");
            htmlStr.Append("			<td class=xl717651 width=41 style='width: 36.81pt'>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("			<td class=xl717651 width=98 style='width: 87.88pt'>&nbsp;</td>                                                                                                                               \n");
            htmlStr.Append("			<td class=xl717651 width=27 style='width: 23.75pt'>&nbsp;</td>                                                                                                                                 \n");
            htmlStr.Append("			<td class=xl717651 width=78 style='width: 68.88pt'>&nbsp;</td>                                                                                                                               \n");
            htmlStr.Append("			<td class=xl1457651 colspan=2>&nbsp;Ngày ký:<span                                                                                                                                           \n");
            htmlStr.Append("				style='mso-spacerun: yes'> </span></td>                                                                                                                                                 \n");
            htmlStr.Append("			<td class=xl1117651>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("			<td colspan=5 class=xl1927651 style='border-right: 0.5pt solid black'>" + dt.Rows[0]["SignedDate"] + "</td>                                                                                                   \n");
            htmlStr.Append("			<td class=xl1097651>&nbsp;</td>                                                                                                                                                             \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
            htmlStr.Append("		<tr height=18 style='mso-height-source: userset; height: 13.95pt'>                                                                                                                              \n");
            htmlStr.Append("			<td height=18 class=xl1377651 style='height: 13.95pt'>&nbsp;</td>                                                                                                                           \n");
            htmlStr.Append("			<td class=xl1147651 colspan=3>Mã CQT: " + dt.Rows[0]["cqt_mccqt_id"] + "</td>                                                                                                                                                                       \n");
            htmlStr.Append("			<td class=xl717651 width=41 style='width: 36.81pt'>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("			<td class=xl717651 width=98 style='width: 87.88pt'>&nbsp;</td>                                                                                                                               \n");
            htmlStr.Append("			<td class=xl717651 width=27 style='width: 23.75pt'>&nbsp;</td>                                                                                                                                 \n");
            htmlStr.Append("			<td class=xl717651 width=78 style='width: 68.88pt'>&nbsp;</td>                                                                                                                               \n");
            htmlStr.Append("			<td class=xl717651 width=56 style='width: 49.88pt'>&nbsp;</td>                                                                                                                               \n");
            htmlStr.Append("			<td class=xl717651 width=30 style='width: 27.32pt'>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("			<td class=xl717651 width=6 style='width: 4.75pt'>&nbsp;</td>                                                                                                                                 \n");
            htmlStr.Append("			<td class=xl717651 width=49 style='width: 43.94pt'>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("			<td class=xl717651 width=48 style='width: 42.75pt'>&nbsp;</td>                                                                                                                               \n");
            htmlStr.Append("			<td class=xl717651 width=6 style='width: 4.75pt'>&nbsp;</td>                                                                                                                                 \n");
            htmlStr.Append("			<td class=xl717651 width=55 style='width: 48.68pt'>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("			<td class=xl717651 width=78 style='width: 68.88pt'>&nbsp;</td>                                                                                                                               \n");
            htmlStr.Append("			<td class=xl1007651 width=6 style='width: 4.75pt'>&nbsp;</td>                                                                                                                                \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
            htmlStr.Append("		<tr height=20 style='mso-height-source: userset; height: 18.75pt'>                                                                                                                               \n");
            htmlStr.Append("			<td height=20 class=xl1377651 style='height: 18.75pt'>&nbsp;</td>                                                                                                                            \n");
            htmlStr.Append("			<td class=xl657652 colspan=7>Tra c&#7913;u t&#7841;i Website: <font                                                                                                                         \n");
            htmlStr.Append("				class='font57651'><span style='mso-spacerun: yes'> </span></font><font                                                                                                                  \n");
            htmlStr.Append("				class='font137651'>" + dt.Rows[0]["WEBSITE_EI"] + "</font></td>                                                                                                                     \n");
            htmlStr.Append("			<td class=xl1147651 colspan=4>Mã nh&#7853;n hóa &#273;&#417;n:                                                                                                                              \n");
            htmlStr.Append("				<font class='font17221001'>" + dt.Rows[0]["matracuu"] + "</font>                                                                                                                                      \n");
            htmlStr.Append("			</td>                                                                                                                                                                                       \n");
            htmlStr.Append("			<td class=xl717651 width=48 style='width: 42.75pt'>&nbsp;</td>                                                                                                                               \n");
            htmlStr.Append("			<td class=xl717651 width=6 style='width: 4.75pt'>&nbsp;</td>                                                                                                                                 \n");
            htmlStr.Append("			<td class=xl717651 width=55 style='width: 48.68pt'>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("			<td class=xl717651 width=78 style='width: 68.88pt'>&nbsp;</td>                                                                                                                               \n");
            htmlStr.Append("			<td class=xl1007651 width=6 style='width: 4.75pt'>&nbsp;</td>                                                                                                                                \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
            htmlStr.Append("		<tr height=16 style='mso-height-source: userset; height: 12.0pt'>                                                                                                                               \n");
            htmlStr.Append("			<td height=16 class=xl827651 style='height: 12.0pt'>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("			<td colspan=15 class=xl1947651>(C&#7847;n ki&#7875;m tra,                                                                                                                                   \n");
            htmlStr.Append("				&#273;&#7889;i chi&#7871;u khi l&#7853;p, giao nh&#7853;n hóa                                                                                                                           \n");
            htmlStr.Append("				&#273;&#417;n)</td>                                                                                                                                                                     \n");
            htmlStr.Append("			<td class=xl1027651 width=6 style='width: 4.75pt'>&nbsp;</td>                                                                                                                                \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
            htmlStr.Append("		<tr height=16 style='mso-height-source: userset; height: 12.0pt'>                                                                                                                               \n");
            htmlStr.Append("			<td height=16 class=xl677651                                                                                                                                                                \n");
            htmlStr.Append("				style='height: 12.0pt; border-top: none'>&nbsp;</td>                                                                                                                                    \n");
            htmlStr.Append("			<td colspan=15 class=xl1957651>" + dt.Rows[0]["CONTRACT_INFO_EI"] + "</td>                                                                                                                                                        \n");
            htmlStr.Append("			<td class=xl1367651 width=6 style='border-top: none; width: 4.75pt'>&nbsp;</td>                                                                                                              \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
            htmlStr.Append("		<![if supportMisalignedColumns]>                                                                                                                                                                \n");
            htmlStr.Append("		<tr height=0 style='display: none'>                                                                                                                                                             \n");
            htmlStr.Append("			<td width=6 style='width: 4.75pt'></td>                                                                                                                                                      \n");
            htmlStr.Append("			<td width=33 style='width: 29.68pt'></td>                                                                                                                                                   \n");
            htmlStr.Append("			<td width=70 style='width: 49.4pt'></td>                                                                                                                                                    \n");
            htmlStr.Append("			<td width=55 style='width: 48.68pt'></td>                                                                                                                                                   \n");
            htmlStr.Append("			<td width=41 style='width: 36.81pt'></td>                                                                                                                                                   \n");
            htmlStr.Append("			<td width=98 style='width: 87.88pt'></td>                                                                                                                                                    \n");
            htmlStr.Append("			<td width=27 style='width: 23.75pt'></td>                                                                                                                                                      \n");
            htmlStr.Append("			<td width=78 style='width: 68.88pt'></td>                                                                                                                                                    \n");
            htmlStr.Append("			<td width=56 style='width: 49.88pt'></td>                                                                                                                                                    \n");
            htmlStr.Append("			<td width=30 style='width: 27.32pt'></td>                                                                                                                                                   \n");
            htmlStr.Append("			<td width=6 style='width: 4.75pt'></td>                                                                                                                                                      \n");
            htmlStr.Append("			<td width=49 style='width: 43.94pt'></td>                                                                                                                                                   \n");
            htmlStr.Append("			<td width=48 style='width: 42.75pt'></td>                                                                                                                                                    \n");
            htmlStr.Append("			<td width=6 style='width: 4.75pt'></td>                                                                                                                                                      \n");
            htmlStr.Append("			<td width=55 style='width: 48.68pt'></td>                                                                                                                                                   \n");
            htmlStr.Append("			<td width=78 style='width: 68.88pt'></td>                                                                                                                                                    \n");
            htmlStr.Append("			<td width=6 style='width: 4.75pt'></td>                                                                                                                                                      \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                           \n");
            htmlStr.Append("		<![endif]>                                                                                                                                                                                      \n");
            htmlStr.Append("	</table>                                                                                                                                                                                            \n");
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
            if (s.Trim().Substring(0, 1) == "-")
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
            if (l == 0)
            {
                rtnf += " Không";
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
        private static int countLength_v2(String s)
        {
            int l_length = 0;//s.length();
            char l_slit;
            int count = 0;
            int v_checking = 0; //l_length/40;
            int result = 0;
            int max_length = 40;
            int index_length = 0;
            //String[] words = s.split("&#xA;");//tach chuoi dua tren khoang trang  &#xA;
            List<string> words = new List<string>(s.Split(new string[] { "&#xA;" }, StringSplitOptions.None));
            result = words.Count;

            //ESysLib.WriteLogError("countLength_v2  result = words.Count; " + result);
            if (result == 1)
            {
                result = countLength(s);
            }
            if (result == 0)
            {
                result = 1;
            }

            return result;
        }
        private static int countLength(string s)
        {
            int result = 0;
            int max_length = 40;
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