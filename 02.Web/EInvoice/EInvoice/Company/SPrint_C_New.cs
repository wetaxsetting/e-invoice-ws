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
    public class SPrint_C_New
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
            v_count = dt_d.Rows.Count;
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


            //string read_prive = "", read_en = "", read_amount = "", amout_vat = "";
            //read_amount = dt.Rows[0]["TotalAmountInWord"].ToString();
            //
            //if (dt.Rows[0]["CurrencyCodeUSD"].ToString() == "VND")
            //{
            //    read_prive = NumberToTextVN(Decimal.Parse(dt.Rows[0]["TotalAmountInWord"].ToString()));
            //}
            //else
            //{
            //    read_prive = Num2VNText(dt.Rows[0]["TotalAmountInWord"].ToString(), "USD");
            //}
            //read_prive = read_prive.Replace(",", "").Replace("TRừ", "Trừ");

            //read_prive = read_prive.ToString().Substring(0, 2) + read_prive.ToString().Substring(2, read_prive.Length - 2).ToLower().Replace("mỹ", "Mỹ");

            //read_prive = read_prive.Substring(0, 2) + read_prive.Substring(2, read_prive.Length - 2).ToLower() + '.';

            read_prive = dt.Rows[0]["amount_word_vie"].ToString();

            //read_en = dt.Rows[0]["amount_word_eng"].ToString();

            if (dt.Rows[0]["vatamount_display"].ToString().Trim() == "0")
            {
                amount_vat = "-";
            }
           

            //read_en = dt.Rows[0]["TotalAmountInWord"].ToString();
            int end = 0;
            int count = count_page_v + r;
            double height = 130;
            StringBuilder htmlStr = new StringBuilder("");
            string heigh = "", heigh_d = "";


            htmlStr.Append("<!DOCTYPE html PUBLIC '-//W3C//DTD HTML 4.01 Transitional//EN' 'http://www.w3.org/TR/html4/loose.dtd'>																													\n");
            htmlStr.Append("<html>                                                                                                                                                                                                                  \n");
            htmlStr.Append("<head>                                                                                                                                                                                                                  \n");
            htmlStr.Append("<meta http-equiv='Content-Type' content='text/html; charset=UTF-8'>                                                                                                                                                     \n");
            htmlStr.Append("<style>                                                                                                                                                                                                                 \n");
            htmlStr.Append("@page {                                                                                                                                                                                                                 \n");
            htmlStr.Append("	size: A4                                                                                                                                                                                                            \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("</style>                                                                                                                                                                                                                \n");
            htmlStr.Append("<link href='https://fonts.googleapis.com/css?family=Tangerine:700'                                                                                                                                                      \n");
            htmlStr.Append("	rel='stylesheet' type='text/css'>                                                                                                                                                                                   \n");
            htmlStr.Append("<style>                                                                                                                                                                                                                 \n");
            htmlStr.Append("/*body   { font-family: serif }                                                                                                                                                                                         \n");
            htmlStr.Append("    h1     { font-family: 'Tangerine', cursive; font-size: 40pt; line-height: 18mm}                                                                                                                                     \n");
            htmlStr.Append("    h2, h3 { font-family: 'Tangerine', cursive; font-size: 24pt; line-height: 7mm }                                                                                                                                     \n");
            htmlStr.Append("    h4     { font-size: 13pt; line-height: 1mm }                                                                                                                                                                        \n");
            htmlStr.Append("    h2 + p { font-size: 18pt; line-height: 7mm }                                                                                                                                                                        \n");
            htmlStr.Append("    h3 + p { font-size: 14pt; line-height: 7mm }                                                                                                                                                                        \n");
            htmlStr.Append("    li     { font-size: 11pt; line-height: 5mm }                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                        \n");
            htmlStr.Append("    h1      { margin: 0 }                                                                                                                                                                                               \n");
            htmlStr.Append("    h1 + ul { margin: 2mm 0 5mm }                                                                                                                                                                                       \n");
            htmlStr.Append("    h2, h3  { margin: 0 3mm 3mm 0; float: left }                                                                                                                                                                        \n");
            htmlStr.Append("    h2 + p,                                                                                                                                                                                                             \n");
            htmlStr.Append("    h3 + p  { margin: 0 0 3mm 50mm }                                                                                                                                                                                    \n");
            htmlStr.Append("    //h4      { margin: 1mm 0 0 2mm; border-bottom: 1px solid black }                                                                                                                                                   \n");
            htmlStr.Append("    h4 + ul { margin: 5mm 0 0 50mm }                                                                                                                                                                                    \n");
            htmlStr.Append("    article { border: 4px double black; padding: 5mm 10mm; border-radius: 3mm }*/                                                                                                                                       \n");
            htmlStr.Append("body {                                                                                                                                                                                                                  \n");
            htmlStr.Append("	color: blue;                                                                                                                                                                                                        \n");
            htmlStr.Append("	font-size: 100%;                                                                                                                                                                                                    \n");
            htmlStr.Append("	background-image: url('assets/Solution.jpg');                                                                                                                                                                       \n");
            htmlStr.Append("}                                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                                        \n");
            htmlStr.Append("h1 {                                                                                                                                                                                                                    \n");
            htmlStr.Append("	color: #00FF00;                                                                                                                                                                                                     \n");
            htmlStr.Append("}                                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                                        \n");
            htmlStr.Append("p {                                                                                                                                                                                                                     \n");
            htmlStr.Append("	color: rgb(0, 0, 255)                                                                                                                                                                                               \n");
            htmlStr.Append("}                                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                                        \n");
            htmlStr.Append("headline1 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	background-image: url(assets/Solution.jpg);                                                                                                                                                                         \n");
            htmlStr.Append("	background-repeat: no-repeat;                                                                                                                                                                                       \n");
            htmlStr.Append("	background-position: left top;                                                                                                                                                                                      \n");
            htmlStr.Append("	padding-top: 68px;                                                                                                                                                                                                  \n");
            htmlStr.Append("	margin-bottom: 50px;                                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                                        \n");
            htmlStr.Append("headline2 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	background-image: url(images/newsletter_headline2.gif);                                                                                                                                                             \n");
            htmlStr.Append("	background-repeat: no-repeat;                                                                                                                                                                                       \n");
            htmlStr.Append("	background-position: left top;                                                                                                                                                                                      \n");
            htmlStr.Append("	padding-top: 68px;                                                                                                                                                                                                  \n");
            htmlStr.Append("}                                                                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                                                        \n");
            htmlStr.Append("<!--table                                                                                                                                                                                                               \n");
            htmlStr.Append("	{mso-displayed-decimal-separator:'\\.';                                                                                                                                                                              \n");
            htmlStr.Append("	mso-displayed-thousand-separator:'\\,';}                                                                                                                                                                             \n");
            htmlStr.Append(".font512067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{color:windowtext;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size:14.4pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                                                                \n");
            htmlStr.Append(".font612067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{color:windowtext;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                                                                \n");
            htmlStr.Append(".font712067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{color:windowtext;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size:14.4pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                                                                \n");
            htmlStr.Append(".font7120671                                                                                                                                                                                                            \n");
            htmlStr.Append("	{color:windowtext;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size:14.4pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	white-space:normal;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                                                                \n");
            htmlStr.Append(".font812067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{color:windowtext;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                                                                \n");
            htmlStr.Append(".font912067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{color:windowtext;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                                                                \n");
            htmlStr.Append(".font1012067                                                                                                                                                                                                            \n");
            htmlStr.Append("	{color:windowtext;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                                                                \n");
            htmlStr.Append(".font1112067                                                                                                                                                                                                            \n");
            htmlStr.Append("	{color:#0066CC;                                                                                                                                                                                                     \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                                                                \n");
            htmlStr.Append(".font1212067                                                                                                                                                                                                            \n");
            htmlStr.Append("	{color:red;                                                                                                                                                                                                         \n");
            htmlStr.Append("	font-size:10.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                                                                \n");
            htmlStr.Append(".font1312067                                                                                                                                                                                                            \n");
            htmlStr.Append("	{color:black;                                                                                                                                                                                                       \n");
            htmlStr.Append("	font-size:14.4pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                                                                \n");
            htmlStr.Append(".font1412067                                                                                                                                                                                                            \n");
            htmlStr.Append("	{color:windowtext;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                                                                \n");
            htmlStr.Append(".font1512067                                                                                                                                                                                                            \n");
            htmlStr.Append("	{color:red;                                                                                                                                                                                                         \n");
            htmlStr.Append("	font-size:14.4pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl6512067                                                                                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl65120671                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:0.6pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl6612067                                                                                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl66120671                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:0.6pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl6712067                                                                                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl67120671                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:0.6pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl6812067                                                                                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:0.6pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl6912067                                                                                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl7012067                                                                                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl70120671                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:0.6pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl7112067                                                                                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:14.4pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl7212067                                                                                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl7312067                                                                                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.44pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl7412067                                                                                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:14.4pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl7512067                                                                                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:14.4pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl7612067                                                                                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl76120671                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl7712067                                                                                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl7812067                                                                                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:15.6pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl7912067                                                                                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl8012067                                                                                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:14.4pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl8112067                                                                                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl8212067                                                                                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl8312067                                                                                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:14.4pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl8412067                                                                                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:14.4pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                                                                                             \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                                                      \n");
            htmlStr.Append(".xl8512067                                                                                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl8612067                                                                                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl8712067                                                                                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl8812067                                                                                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl8912067                                                                                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:14.4pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl9012067                                                                                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl9112067                                                                                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl9212067                                                                                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl9312067                                                                                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:14.4pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl9412067                                                                                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl9512067                                                                                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:14.4pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl9612067                                                                                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:#C00000;                                                                                                                                                                                                      \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl9712067                                                                                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:14.4pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl9812067                                                                                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl9912067                                                                                                                                                                                                              \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:17.28pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl10012067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:17.28pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl10112067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:15.6pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl10212067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:15.6pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl10312067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:15.6pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl10412067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:red;                                                                                                                                                                                                          \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl10512067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:15.6pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl10612067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:red;                                                                                                                                                                                                          \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl10712067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:#0070C0;                                                                                                                                                                                                      \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl10812067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                                                      \n");
            htmlStr.Append(".xl10912067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                                                      \n");
            htmlStr.Append(".xl11012067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:14.4pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                                                      \n");
            htmlStr.Append(".xl11112067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:red;                                                                                                                                                                                                          \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                                                      \n");
            htmlStr.Append(".xl11212067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:17.28pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl11312067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl11412067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                                                                                             \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:1.0pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:1.0pt dotted windowtext;                                                                                                                                                                               \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                                                      \n");
            htmlStr.Append(".xl11512067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                                                                                             \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:1.0pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:1.0pt dotted windowtext;                                                                                                                                                                               \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                                                      \n");
            htmlStr.Append(".xl11612067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                                                                                             \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:1.0pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl11712067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:14.4pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                                                                                             \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                                                      \n");
            htmlStr.Append(".xl11812067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:1.0pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:1.0pt dotted windowtext;                                                                                                                                                                               \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                                                      \n");
            htmlStr.Append(".xl11912067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                                                                                             \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:1.0pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:1.0pt dotted windowtext;                                                                                                                                                                               \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                                                      \n");
            htmlStr.Append(".xl12012067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:1.0pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl12112067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                                                                                             \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:1.0pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl12212067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                                                                                             \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:1.0pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl12312067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                                                                                             \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:1.0pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:1.0pt dotted windowtext;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl12412067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                                                                                             \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:1.0pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl12512067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl12612067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                                                                                             \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:1.0pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl12712067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:0%;                                                                                                                                                                                               \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl12812067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:14.4pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl12912067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:black;                                                                                                                                                                                                        \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl13012067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:black;                                                                                                                                                                                                        \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl13112067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:14.4pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl13212067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl13312067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:1.0pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:1.0pt dotted windowtext;                                                                                                                                                                               \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl13412067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:1.0pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:1.0pt dotted windowtext;                                                                                                                                                                               \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl13512067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:14.4pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:1.0pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:1.0pt dotted windowtext;                                                                                                                                                                               \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl13612067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:14.4pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:1.0pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:1.0pt dotted windowtext;                                                                                                                                                                               \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl13712067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:14.4pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:1.0pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:1.0pt dotted windowtext;                                                                                                                                                                               \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl13812067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                                                                                             \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:1.0pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:1.0pt dotted windowtext;                                                                                                                                                                               \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                                                      \n");
            htmlStr.Append(".xl13912067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                                                                                             \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:1.0pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:1.0pt dotted windowtext;                                                                                                                                                                               \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                                                      \n");
            htmlStr.Append(".xl14012067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:1.0pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:1.0pt dotted windowtext;                                                                                                                                                                               \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl14112067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:1.0pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:1.0pt dotted windowtext;                                                                                                                                                                               \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl14212067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl14312067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                                                      \n");
            htmlStr.Append(".xl14412067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:14.4pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl14512067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl145120671                                                                                                                                                                                                            \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:0.6pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl14612067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:#002060;                                                                                                                                                                                                      \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                                                      \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source:auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern:auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl14712067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:#002060;                                                                                                                                                                                                      \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                                                      \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	mso-background-source:auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern:auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl14812067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:#002060;                                                                                                                                                                                                      \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                                                      \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	mso-background-source:auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern:auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl14912067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:#002060;                                                                                                                                                                                                      \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                                                      \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	mso-background-source:auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern:auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl15012067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:#002060;                                                                                                                                                                                                      \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                                                      \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	mso-background-source:auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern:auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl15112067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:#002060;                                                                                                                                                                                                      \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                                                      \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	mso-background-source:auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern:auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl15212067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl15312067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:red;                                                                                                                                                                                                          \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:'Short Date';                                                                                                                                                                                     \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl15412067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:red;                                                                                                                                                                                                          \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:'Short Date';                                                                                                                                                                                     \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl15512067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl15612067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl15712067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:15.36pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl15812067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:15.36pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl15912067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:14.4pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl16012067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:red;                                                                                                                                                                                                          \n");
            htmlStr.Append("	font-size:14.4pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl16112067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:14.4pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl16212067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl16312067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:14.4pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl163120671                                                                                                                                                                                                            \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl16412067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:#002060;                                                                                                                                                                                                      \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                                                      \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source:auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern:auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                                                      \n");
            htmlStr.Append(".xl16512067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:red;                                                                                                                                                                                                          \n");
            htmlStr.Append("	font-size:16.32pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl16612067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:red;                                                                                                                                                                                                          \n");
            htmlStr.Append("	font-size:13.44pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl16712067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:14.4pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl16812067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                                                      \n");
            htmlStr.Append(".xl16912067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                                                      \n");
            htmlStr.Append(".xl17012067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                                                      \n");
            htmlStr.Append(".xl17112067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl17212067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:14.4pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl17312067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:14.4pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl17412067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl17512067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl17612067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl17712067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl17812067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl17912067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl18012067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl18112067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                                                                                             \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:1.0pt dotted windowtext;                                                                                                                                                                               \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                                                      \n");
            htmlStr.Append(".xl18212067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                                                                                             \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:1.0pt dotted windowtext;                                                                                                                                                                               \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                                                      \n");
            htmlStr.Append(".xl18312067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:red;                                                                                                                                                                                                          \n");
            htmlStr.Append("	font-size:14.4pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:1.0pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:1.0pt dotted windowtext;                                                                                                                                                                               \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl18412067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:red;                                                                                                                                                                                                          \n");
            htmlStr.Append("	font-size:14.4pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:1.0pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:1.0pt dotted windowtext;                                                                                                                                                                               \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl18512067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:red;                                                                                                                                                                                                          \n");
            htmlStr.Append("	font-size:14.4pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:1.0pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:1.0pt dotted windowtext;                                                                                                                                                                               \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl18612067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl18712067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl18812067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                                                                                             \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                                                      \n");
            htmlStr.Append(".xl18912067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                                                                                             \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                                                      \n");
            htmlStr.Append(".xl19012067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:14.4pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl19112067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:1.0pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl19212067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:1.0pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl19312067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:14.4pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:1.0pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl19412067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                                                                                             \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:1.0pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:1.0pt dotted windowtext;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                                                      \n");
            htmlStr.Append(".xl19512067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                                                                                             \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:1.0pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                                                      \n");
            htmlStr.Append(".xl19612067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:.5pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl19712067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:12.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl19812067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl19912067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:14.4pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl20012067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append(".xl20112067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:black;                                                                                                                                                                                                        \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:.5pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                                                      \n");
            htmlStr.Append(".xl20212067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:black;                                                                                                                                                                                                        \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                                                      \n");
            htmlStr.Append(".xl20312067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:black;                                                                                                                                                                                                        \n");
            htmlStr.Append("	font-size:13.2pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right:.5pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                                                      \n");
            htmlStr.Append(".xl20412067                                                                                                                                                                                                             \n");
            htmlStr.Append("	{padding:0px;                                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:14.4pt;                                                                                                                                                                                                  \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top:.5pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	background:white;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-pattern:black none;                                                                                                                                                                                             \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                                                \n");
            htmlStr.Append("-->                                                                                                                                                                                                                     \n");
            htmlStr.Append("</style>                                                                                                                                                                                                                \n");
            htmlStr.Append("</head>                                                                                                                                                                                                                 \n");
            htmlStr.Append("<body class='A4'>                                                                                                                                                                                                       \n");

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

            double v_totalHeightLastPage = 203.5;// 243.5.5;

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

                htmlStr.Append("<table border=0 cellpadding=0 cellspacing=0 width=746 class=xl6512067                                                                                                                                                   \n");
                htmlStr.Append(" style='border-collapse:collapse;table-layout:fixed;width:667.5pt'>                                                                                                                                                       \n");
                htmlStr.Append(" <col class=xl6512067 width=6 style='mso-width-source:userset;mso-width-alt:                                                                                                                                            \n");
                htmlStr.Append(" 199;width:3.84pt'>                                                                                                                                                                                                     \n");
                htmlStr.Append(" <col class=xl6512067 width=33 style='mso-width-source:userset;mso-width-alt:                                                                                                                                           \n");
                htmlStr.Append(" 1166;width:24pt'>                                                                                                                                                                                                      \n");
                htmlStr.Append(" <col class=xl6512067 width=70 style='mso-width-source:userset;mso-width-alt:                                                                                                                                           \n");
                htmlStr.Append(" 2474;width:59.92pt'>                                                                                                                                                                                                   \n");
                htmlStr.Append(" <col class=xl6512067 width=55 style='mso-width-source:userset;mso-width-alt:                                                                                                                                           \n");
                htmlStr.Append(" 1962;width:29.36pt'>                                                                                                                                                                                                   \n");
                htmlStr.Append(" <col class=xl6512067 width=41 span=2 style='mso-width-source:userset;                                                                                                                                                  \n");
                htmlStr.Append(" mso-width-alt:1450;width:20.76pt'><!--width:29.76pt  -->                                                                                                                                                               \n");
                htmlStr.Append(" <col class=xl6512067 width=70 style='mso-width-source:userset;mso-width-alt:                                                                                                                                           \n");
                htmlStr.Append(" 2474;width:58.92pt'>                                                                                                                                                                                                   \n");
                htmlStr.Append(" <col class=xl6512067 width=12 style='mso-width-source:userset;mso-width-alt:                                                                                                                                           \n");
                htmlStr.Append(" 426;width:8.64pt'>                                                                                                                                                                                                     \n");
                htmlStr.Append(" <col class=xl6512067 width=78 style='mso-width-source:userset;mso-width-alt:                                                                                                                                           \n");
                htmlStr.Append(" 2759;width:55.68pt'>                                                                                                                                                                                                   \n");
                htmlStr.Append(" <col class=xl6512067 width=56 style='mso-width-source:userset;mso-width-alt:                                                                                                                                           \n");
                htmlStr.Append(" 1991;width:40.32pt'>                                                                                                                                                                                                   \n");
                htmlStr.Append(" <col class=xl6512067 width=30 style='mso-width-source:userset;mso-width-alt:                                                                                                                                           \n");
                htmlStr.Append(" 1080;width:22.08pt'>                                                                                                                                                                                                   \n");
                htmlStr.Append(" <col class=xl6512067 width=6 style='mso-width-source:userset;mso-width-alt:                                                                                                                                            \n");
                htmlStr.Append(" 199;width:3.84pt'>                                                                                                                                                                                                     \n");
                htmlStr.Append(" <col class=xl6512067 width=41 style='mso-width-source:userset;mso-width-alt:                                                                                                                                           \n");
                htmlStr.Append(" 1450;width:29.76pt'>                                                                                                                                                                                                   \n");
                htmlStr.Append(" <col class=xl6512067 width=48 style='mso-width-source:userset;mso-width-alt:                                                                                                                                           \n");
                htmlStr.Append(" 1706;width:34.56pt'>                                                                                                                                                                                                   \n");
                htmlStr.Append(" <col class=xl6512067 width=6 style='mso-width-source:userset;mso-width-alt:                                                                                                                                            \n");
                htmlStr.Append(" 199;width:3.84pt'>                                                                                                                                                                                                     \n");
                htmlStr.Append(" <col class=xl6512067 width=41 style='mso-width-source:userset;mso-width-alt:                                                                                                                                           \n");
                htmlStr.Append(" 1450;width:29.76pt'>                                                                                                                                                                                                   \n");
                htmlStr.Append(" <col class=xl6512067 width=106 style='mso-width-source:userset;mso-width-alt:                                                                                                                                          \n");
                htmlStr.Append(" 3754;width:75.84pt'>                                                                                                                                                                                                    \n");
                htmlStr.Append(" <col class=xl6512067 width=6 style='mso-width-source:userset;mso-width-alt:                                                                                                                                            \n");
                htmlStr.Append(" 199;width:3.84pt'>                                                                                                                                                                                                     \n");
                htmlStr.Append(" <tr height=33 style='mso-height-source:userset;height:30.06pt'>																																						\n");
                htmlStr.Append("  <td height=33 class=xl11212067 width=6 style='height:30.06pt;width:3.84pt'>&nbsp;</td>                                                                                                                               \n");
                htmlStr.Append("  <td width=33 class=xl6712067 style='width:24pt' align=left valign=top><![if !vml]><span style='mso-ignore:vglayout;                                                                                                                   \n");
                htmlStr.Append("  position:absolute;z-index:2;margin-left:6px;margin-top:7px;width:147.5px;                                                                                                                                               \n");
                htmlStr.Append("  height:103.75px'><img width=147.5 height=103.75                                                                                                                                                                                 \n");
                htmlStr.Append("  src='D:/webproject/e-invoice-ws/02.Web/EInvoice/img/SPRINT_001.png'                                                                                                                                                 \n");
                htmlStr.Append("  alt='Description: cid:image001.jpg@01D270B2.087E24E0' v:shapes='Picture_x0020_5'></span><![endif]><span                                                                                                               \n");
                htmlStr.Append("  style='mso-ignore:vglayout2'>                                                                                                                                                                                         \n");
                htmlStr.Append("  <table cellpadding=0 cellspacing=0>                                                                                                                                                                                   \n");
                htmlStr.Append("   <tr>                                                                                                                                                                                                                 \n");
                htmlStr.Append("    <td height=33  width=33 style='height:30.06pt;width:24pt'>&nbsp;</td>                                                                                                                               \n");
                htmlStr.Append("   </tr>                                                                                                                                                                                                                \n");
                htmlStr.Append("  </table>                                                                                                                                                                                                              \n");
                htmlStr.Append("  </span></td>                                                                                                                                                                                                          \n");
                htmlStr.Append("  <td class=xl9912067 width=70 style='width:49.92pt'>&nbsp;</td>                                                                                                                                                        \n");
                htmlStr.Append("  <td class=xl9912067 width=55 style='width:39.36pt'>&nbsp;</td>                                                                                                                                                        \n");
                htmlStr.Append("  <td class=xl15812067 colspan=2 width=82 style='width:62pt'>" + dt.Rows[0]["Seller_Name"] + "</td>                                                                                                                                    \n");
                htmlStr.Append("  <td class=xl14612067 width=70 style='width:49.92pt'>&nbsp;</td>                                                                                                                                                       \n");
                htmlStr.Append("  <td class=xl14612067 width=12 style='width:8.64pt'>&nbsp;</td>                                                                                                                                                        \n");
                htmlStr.Append("  <td class=xl14612067 width=78 style='width:55.68pt'>&nbsp;</td>                                                                                                                                                       \n");
                htmlStr.Append("  <td class=xl14612067 width=56 style='width:40.32pt'>&nbsp;</td>                                                                                                                                                       \n");
                htmlStr.Append("  <td class=xl14612067 width=30 style='width:22.08pt'>&nbsp;</td>                                                                                                                                                       \n");
                htmlStr.Append("  <td class=xl14612067 width=6 style='width:3.84pt'>&nbsp;</td>                                                                                                                                                         \n");
                htmlStr.Append("  <td class=xl14612067 width=41 style='width:29.76pt'>&nbsp;</td>                                                                                                                                                       \n");
                htmlStr.Append("  <td class=xl9912067 width=48 style='width:34.56pt'>&nbsp;</td>                                                                                                                                                        \n");
                htmlStr.Append("  <td class=xl9912067 width=6 style='width:3.84pt'>&nbsp;</td>                                                                                                                                                          \n");
                htmlStr.Append("  <td class=xl9912067 width=41 style='width:29.76pt'>&nbsp;</td>                                                                                                                                                        \n");
                htmlStr.Append("  <td class=xl9912067 width=106 style='width75.84pt'>&nbsp;</td>                                                                                                                                                        \n");
                htmlStr.Append("  <td class=xl10012067 width=6 style='width:3.84pt'>&nbsp;</td>                                                                                                                                                         \n");
                htmlStr.Append(" </tr>                                                                                                                                                                                                                  \n");
                htmlStr.Append(" <tr height=24 style='mso-height-source:userset;height:23.18pt'>                                                                                                                                                        \n");
                htmlStr.Append("  <td height=24 class=xl14512067 style='height:23.18pt'>&nbsp;</td>                                                                                                                                                     \n");
                htmlStr.Append("  <td class=xl6512067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append("  <td class=xl6512067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append("  <td class=xl6512067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append("  <td class=xl7412067 colspan=2>Mã s&#7889; thu&#7871; (<font class='font1012067'>Tax code</font><font class='font712067'>)</font></td>                                                                                 \n");
                htmlStr.Append("  <td class=xl14712067></td>                                                                                                                                                                                            \n");
                htmlStr.Append("  <td class=xl14912067>:</td>                                                                                                                                                                                           \n");
                htmlStr.Append("  <td colspan=9 class=xl16112067 width=412 style='width:308pt'>" + dt.Rows[0]["Seller_TaxCode"] + "</td>                                                                                                                             \n");
                htmlStr.Append("  <td class=xl7012067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append(" </tr>                                                                                                                                                                                                                  \n");
                htmlStr.Append(" <tr height=33 style='mso-height-source:userset;height:30.06pt'>                                                                                                                                                       \n");
                htmlStr.Append("  <td height=33 class=xl14512067 style='height:30.06pt'>&nbsp;</td>                                                                                                                                                    \n");
                htmlStr.Append("  <td class=xl6512067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append("  <td class=xl6512067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append("  <td class=xl6512067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append("  <td class=xl7412067 colspan=2>&#272;&#7883;a ch&#7881; (<font                                                                                                                                                         \n");
                htmlStr.Append("  class='font1012067'>Address</font><font class='font712067'>)</font></td>                                                                                                                                              \n");
                htmlStr.Append("  <td class=xl14812067 width=70 style='width:49.92pt'></td>                                                                                                                                                             \n");
                htmlStr.Append("  <td class=xl15012067 width=12 style='width:8.64pt'>:</td>                                                                                                                                                             \n");
                htmlStr.Append("  <td colspan=9 class=xl16212067 width=412 style='width:308pt'>" + dt.Rows[0]["Seller_Address"] + "</td>                                                                                                                             \n");
                htmlStr.Append("  <td class=xl7012067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append(" </tr>                                                                                                                                                                                                                  \n");
                htmlStr.Append(" <tr height=20 style='mso-height-source:userset;height:23.18pt'>                                                                                                                                                        \n");
                htmlStr.Append("  <td height=20 class=xl14512067 style='height:23.18pt'>&nbsp;</td>                                                                                                                                                     \n");
                htmlStr.Append("  <td class=xl6512067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append("  <td class=xl6512067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append("  <td class=xl6512067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append("  <td class=xl7412067 colspan=2>&#272;i&#7879;n tho&#7841;i (<font class='font912067'>Tel)</font></td>                                                                                                                  \n");
                htmlStr.Append("  <td class=xl14712067 dir=LTR></td>                                                                                                                                                                                    \n");
                htmlStr.Append("  <td class=xl15112067>:</td>                                                                                                                                                                                           \n");
                htmlStr.Append("  <td colspan=5 class=xl163120671 width=211 style='width:158pt'>" + dt.Rows[0]["Seller_Tel"] + "</td>                                                                                                                                \n");
                htmlStr.Append("  <td class=xl7412067 colspan=4>Fax: " + dt.Rows[0]["Seller_Fax"] + "</td>                                                                                                                                                           \n");
                htmlStr.Append("  <td class=xl7212067 width=6 style='width:3.84pt'>&nbsp;</td>                                                                                                                                                          \n");
                htmlStr.Append(" </tr>                                                                                                                                                                                                                  \n");
                htmlStr.Append(" <tr height=21 style='height:23.18pt'>                                                                                                                                                                                  \n");
                htmlStr.Append("  <td height=21 class=xl14512067 style='height:23.18pt'>&nbsp;</td>                                                                                                                                                     \n");
                htmlStr.Append("  <td class=xl6512067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append("  <td class=xl6512067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append("  <td class=xl6512067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append("  <td class=xl7412067 colspan=2>S&#7889; tài kho&#7843;n (<font class='font1012067'>Acc. code</font><font class='font712067'>)</font></td>                                                                              \n");
                htmlStr.Append("  <td class=xl14712067 dir=LTR></td>                                                                                                                                                                                    \n");
                htmlStr.Append("  <td class=xl15112067>:</td>                                                                                                                                                                                           \n");
                htmlStr.Append("  <td colspan=9 class=xl163120671 width=412 style='width:308pt'>" + dt.Rows[0]["SELLER_ACCOUNTNO"] + "  " + dt.Rows[0]["SELLER_ACCOUNTNO2"] + "</td>                                                                                                        \n");
                htmlStr.Append("  <td class=xl7212067 width=6 style='width:3.84pt'>&nbsp;</td>                                                                                                                                                          \n");
                htmlStr.Append(" </tr>                                                                                                                                                                                                                  \n");
                htmlStr.Append(" <tr height=21 style='height:23.18pt'>                                                                                                                                                                                  \n");
                htmlStr.Append("  <td height=21 class=xl14512067 style='height:23.18pt'>&nbsp;</td>                                                                                                                                                     \n");
                htmlStr.Append("  <td class=xl6512067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append("  <td class=xl6512067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append("  <td class=xl6512067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append("  <td class=xl7412067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append("  <td class=xl7412067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append("  <td class=xl14712067 dir=LTR></td>                                                                                                                                                                                    \n");
                htmlStr.Append("  <td class=xl14712067 dir=LTR></td>                                                                                                                                                                                    \n");
                htmlStr.Append("  <td colspan=9 class=xl163120671 width=412 style='width:308pt'>" + dt.Rows[0]["seller_bankname"] + "</td>                                                                                                                                      \n");
                htmlStr.Append("  <td class=xl7212067 width=6 style='width:3.84pt'>&nbsp;</td>                                                                                                                                                          \n");
                htmlStr.Append(" </tr>                                                                                                                                                                                                                  \n");
                htmlStr.Append(" <tr height=4 style='mso-height-source:userset;height:3.0pt'>                                                                                                                                                           \n");
                htmlStr.Append("  <td height=4 class=xl66120671 style='height:3.0pt'>&nbsp;</td>                                                                                                                                                        \n");
                htmlStr.Append("  <td class=xl67120671 colspan=16>&nbsp;</td>                                                                                                                                                                           \n");
                htmlStr.Append("  <td class=xl6812067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append(" </tr>                                                                                                                                                                                                                  \n");
                htmlStr.Append(" <tr height=24 style='mso-height-source:userset;height:20.73pt'>                                                                                                                                                        \n");
                htmlStr.Append("  <td height=24 class=xl14512067 style='height:20.73pt'>&nbsp;</td>                                                                                                                                                     \n");
                htmlStr.Append("  <td class=xl7412067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append("  <td class=xl7412067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append("  <td colspan=10 class=xl16512067>HÓA &#272;&#416;N GIÁ TR&#7882; GIA T&#258;NG</td>                                                                                                                                    \n");
                htmlStr.Append("  <td class=xl6912067 colspan=3><font class='font1012067'></font><font                                                                                                                          \n");
                htmlStr.Append("  class='font812067'></font></td>                                                                                                                                                                                      \n");
                htmlStr.Append("  <td class=xl15912067 colspan=2 style='border-right:.5pt solid black'>                                                                                                                                                \n");
                htmlStr.Append("  </td>                                                                                                                                                                                                  \n");
                htmlStr.Append(" </tr>                                                                                                                                                                                                                  \n");
                htmlStr.Append(" <tr height=24 style='mso-height-source:userset;height:20.73pt'>                                                                                                                                                        \n");
                htmlStr.Append("  <td height=24 class=xl14512067 style='height:20.73pt'>&nbsp;</td>                                                                                                                                                     \n");
                htmlStr.Append("  <td class=xl7412067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append("  <td class=xl7412067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append("  <td colspan=10 class=xl16612067>(VAT INVOICE)</td>                                                                                                                                                                    \n");
                htmlStr.Append("  <td class=xl6912067 colspan=3>Ký hi&#7879;u (<font class='font1012067'>Serial)</font></td>                                                                                                                            \n");
                htmlStr.Append("  <td class=xl15912067>: " + dt.Rows[0]["templateCode"] + "" + dt.Rows[0]["InvoiceSerialNo"] + "</td>                                                                                                                                                                         \n");
                htmlStr.Append("  <td class=xl7012067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append(" </tr>                                                                                                                                                                                                                  \n");
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
                htmlStr.Append(" <tr height=24 style='mso-height-source:userset;height:20.73pt'>                                                                                                                                                        \n");
                htmlStr.Append("  <td height=24 class=xl14512067 style='height:20.73pt'>&nbsp;</td>                                                                                                                                                     \n");
                htmlStr.Append("  <td class=xl7412067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append("  <td class=xl7412067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append("  <td colspan=10 class=xl16712067 width=430 style='width:322pt'>Ngày <font                                                                                                                                              \n");
                htmlStr.Append("  class='font1012067'>(Date)</font><font class='font912067'>&nbsp;" + dt.Rows[0]["invoiceissueddate_dd"] + "&nbsp;<span                                                                                                                              \n");
                htmlStr.Append("  style='mso-spacerun:yes'>   </span></font><font class='font712067'><span                                                                                                                                              \n");
                htmlStr.Append("  style='mso-spacerun:yes'> </span>tháng </font><font class='font1012067'>(month)</font><font                                                                                                                           \n");
                htmlStr.Append("  class='font712067'>&nbsp;" + dt.Rows[0]["invoiceissueddate_mm"] + "&nbsp;<span style='mso-spacerun:yes'>    </span>n&#259;m </font><font                                                                                                           \n");
                htmlStr.Append("  class='font1012067'>(year)</font><font class='font712067'>&nbsp;" + dt.Rows[0]["invoiceissueddate_yyyy"] + "&nbsp;<span                                                                                                                            \n");
                htmlStr.Append("  style='mso-spacerun:yes'>    </span></font></td>                                                                                                                                                                      \n");
                htmlStr.Append("  <td class=xl6512067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append("  <td class=xl6512067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append("  <td class=xl6512067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append("  <td class=xl6512067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append("  <td class=xl7012067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append(" </tr>                                                                                                                                                                                                                  \n");
                htmlStr.Append(" <tr height=6 style='mso-height-source:userset;height:4.95pt'>                                                                                                                                                          \n");
                htmlStr.Append("  <td colspan=2 height=6 class=xl145120671 style='height:4.95pt'>&nbsp;</td>                                                                                                                                            \n");
                htmlStr.Append("  <td colspan=15 class=xl65120671 style='border-bottom:.5pt solid windowtext'>&nbsp;</td>                                                                                                                               \n");
                htmlStr.Append("  <td class=xl70120671>&nbsp;</td>                                                                                                                                                                                      \n");
                htmlStr.Append(" </tr>                                                                                                                                                                                                                  \n");
                htmlStr.Append(" <tr height=4 style='mso-height-source:userset;height:3.0pt'>                                                                                                                                                           \n");
                htmlStr.Append("  <td height=4 class=xl66120671 style='height:3.0pt'>&nbsp;</td>                                                                                                                                                        \n");
                htmlStr.Append("  <td class=xl67120671 colspan=16>&nbsp;</td>                                                                                                                                                                           \n");
                htmlStr.Append("  <td class=xl6812067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append(" </tr>                                                                                                                                                                                                                  \n");
                htmlStr.Append(" <tr height=24 style='mso-height-source:userset;height:20.73pt'>                                                                                                                                                        \n");
                htmlStr.Append("  <td height=24 class=xl14512067 style='height:20.73pt'>&nbsp;</td>                                                                                                                                                     \n");
                htmlStr.Append("  <td class=xl7412067 colspan=6>H&#7885; tên ng&#432;&#7901;i mua hàng (<font                                                                                                                                           \n");
                htmlStr.Append("  class='font1012067'>Customer's name</font><font class='font712067'>):</font></td>                                                                                                                                     \n");
                htmlStr.Append("  <td colspan=11 class=xl17012067 style='border-right:.5pt solid black'>&nbsp;" + dt.Rows[0]["buyer"] + "</td>                                                                                                                   \n");
                htmlStr.Append(" </tr>                                                                                                                                                                                                                  \n");
                htmlStr.Append(" <tr height=29 style='mso-height-source:userset;height:27.57pt'>                                                                                                                                                        \n");
                htmlStr.Append("  <td height=29 class=xl14512067 style='height:27.57pt'>&nbsp;</td>                                                                                                                                                     \n");
                htmlStr.Append("  <td class=xl7412067 colspan=4>Tên &#273;&#417;n v&#7883; (<font                                                                                                                                                       \n");
                htmlStr.Append("  class='font1012067'>Company's name</font><font class='font712067'>):</font></td>                                                                                                                                      \n");
                htmlStr.Append("  <td colspan=12 class=xl17112067>&nbsp;" + dt.Rows[0]["buyerlegalname"] + "</td>                                                                                                                                                         \n");
                htmlStr.Append("  <td class=xl14312067>&nbsp;</td>                                                                                                                                                                                      \n");
                htmlStr.Append(" </tr>                                                                                                                                                                                                                  \n");
                htmlStr.Append(" <tr height=26 style='mso-height-source:userset;height:20.73pt'>                                                                                                                                                        \n");
                htmlStr.Append("  <td height=26 class=xl14512067 style='height:20.73pt'>&nbsp;</td>                                                                                                                                                     \n");
                htmlStr.Append("  <td class=xl7412067 colspan=2>&#272;&#7883;a ch&#7881; (<font                                                                                                                                                         \n");
                htmlStr.Append("  class='font1012067'>Address</font><font class='font712067'>):<span                                                                                                                                                    \n");
                htmlStr.Append("  style='mso-spacerun:yes'> </span></font></td>                                                                                                                                                                         \n");
                htmlStr.Append("  <td class=xl76120671 colspan=14>&nbsp;&nbsp;" + dt.Rows[0]["BuyerAddress"] + "</td>                                                                                                                                                      \n");
                htmlStr.Append("  <td class=xl7712067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append(" </tr>                                                                                                                                                                                                                  \n");
                htmlStr.Append(" <tr height=24 style='mso-height-source:userset;height:20.73pt'>                                                                                                                                                        \n");
                htmlStr.Append("  <td height=24 class=xl14512067 style='height:20.73pt'>&nbsp;</td>                                                                                                                                                     \n");
                htmlStr.Append("  <td class=xl7412067 colspan=3>Mã s&#7889; thu&#7871; (<font                                                                                                                                                           \n");
                htmlStr.Append("  class='font1012067'>Tax code</font><font class='font712067'>):<span                                                                                                                                                   \n");
                htmlStr.Append("  style='mso-spacerun:yes'> </span></font></td>                                                                                                                                                                         \n");
                htmlStr.Append("  <td class=xl76120671>&nbsp;" + dt.Rows[0]["BuyerTaxCode"] + "</td>                                                                                                                                                                      \n");
                htmlStr.Append("  <td class=xl76120671>&nbsp;</td>                                                                                                                                                                                      \n");
                htmlStr.Append("  <td class=xl11312067>&nbsp;</td>                                                                                                                                                                                      \n");
                htmlStr.Append("  <td class=xl11312067>&nbsp;</td>                                                                                                                                                                                      \n");
                htmlStr.Append("  <td class=xl11312067>&nbsp;</td>                                                                                                                                                                                      \n");
                htmlStr.Append("  <td class=xl7412067 colspan=8>Hình th&#7913;c thanh toán (<font                                                                                                                                                       \n");
                htmlStr.Append("  class='font1012067'>Payment Method)</font><font class='xl76120671'>:<span                                                                                                                                             \n");
                htmlStr.Append("  style='mso-spacerun:yes'> </span>" + dt.Rows[0]["PaymentMethodCK"] + "</font></td>                                                                                                                                                   \n");
                htmlStr.Append("  <td class=xl14312067>&nbsp;</td>                                                                                                                                                                                      \n");
                htmlStr.Append(" </tr>                                                                                                                                                                                                                  \n");
                htmlStr.Append(" <tr height=26 style='mso-height-source:userset;height:20.73pt'>                                                                                                                                                        \n");
                htmlStr.Append("  <td height=26 class=xl14512067 style='height:20.73pt'>&nbsp;</td>                                                                                                                                                     \n");
                htmlStr.Append("  <td class=xl6512067 colspan=4><font class='font712067'>Đơn vị tiền tệ (</font><font class='font1012067'>Currency</font><font                                                                                                                                                \n");
                htmlStr.Append("  class='font712067'>):</font> </td>                                                                                                                                                                                     \n");
                htmlStr.Append("  <td class=xl6512067>&nbsp;" + dt.Rows[0]["CurrencyCodeUSD"] + "</td>                                                                                                                                                                                       \n");
                htmlStr.Append("  <td class=xl6512067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append("  <td class=xl7812067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append("  <td class=xl7412067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append("  <td class=xl6512067 colspan=4><font class='font712067'>Tỷ giá (</font><font                                                                                                                              \n");
                htmlStr.Append("  class='font1012067'>Exchange rate</font><font class='font712067'>):</font></td>                                                                                                                                                  \n");
                htmlStr.Append("  <td class=xl6512067>&nbsp;" + dt.Rows[0]["TR_RATE_88"] + "</td>                                                                                                                                                                                       \n");
                htmlStr.Append("  <td class=xl6512067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append("  <td class=xl6512067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append("  <td class=xl6512067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append("  <td class=xl7012067>&nbsp;</td>                                                                                                                                                                                       \n");
                htmlStr.Append(" </tr>                                                                                                                                                                                                                  \n");
                htmlStr.Append(" <tr height=6 style='mso-height-source:userset;height:4.95pt'>                                                                                                                                                          \n");
                htmlStr.Append("  <td colspan=2 height=6 class=xl145120671 style='height:4.95pt'>&nbsp;</td>                                                                                                                                            \n");
                htmlStr.Append("  <td colspan=15 class=xl65120671 style='border-bottom:.5pt solid windowtext'>&nbsp;</td>                                                                                                                               \n");
                htmlStr.Append("  <td class=xl70120671>&nbsp;</td>                                                                                                                                                                                      \n");
                htmlStr.Append(" </tr>                                                                                                                                                                                                                  \n");
                htmlStr.Append("                                                                                                                                                                                                                        \n");
                htmlStr.Append(" <tr class=xl8012067 height=21 style='height:23.18pt'>                                                                                                                                                                  \n");
                htmlStr.Append("  <td colspan=2 height=21 class=xl17212067 style='height:23.18pt'>STT</td>                                                                                                                                              \n");
                htmlStr.Append("  <td colspan=6 class=xl17212067 style='border-right:.5pt solid black'>Tên hàng                                                                                                                                         \n");
                htmlStr.Append("  hóa, d&#7883;ch v&#7909;</td>                                                                                                                                                                                         \n");
                htmlStr.Append("  <td class=xl14412067>&#272;&#417;n v&#7883; tính</td>                                                                                                                                                                 \n");
                htmlStr.Append("  <td colspan=3 class=xl17212067 style='border-right:.5pt solid black'>S&#7889;                                                                                                                                         \n");
                htmlStr.Append("  l&#432;&#7907;ng</td>                                                                                                                                                                                                 \n");
                htmlStr.Append("  <td colspan=3 class=xl14412067>&#272;&#417;n giá</td>                                                                                                                                                                 \n");
                htmlStr.Append("  <td colspan=3 class=xl17212067 style='border-right:.5pt solid black'>Thành                                                                                                                                            \n");
                htmlStr.Append("  ti&#7873;n</td>                                                                                                                                                                                                       \n");
                htmlStr.Append(" </tr>                                                                                                                                                                                                                  \n");
                htmlStr.Append(" <tr class=xl8112067 height=18 style='height:16.5pt'>                                                                                                                                                                   \n");
                htmlStr.Append("  <td colspan=2 height=18 class=xl17412067 style='height:16.5pt'>No.</td>                                                                                                                                               \n");
                htmlStr.Append("  <td colspan=6 class=xl17512067 style='border-right:.5pt solid black'>Name of                                                                                                                                          \n");
                htmlStr.Append("  goods and services</td>                                                                                                                                                                                               \n");
                htmlStr.Append("  <td class=xl8112067>Unit</td>                                                                                                                                                                                         \n");
                htmlStr.Append("  <td colspan=3 class=xl17412067 style='border-right:.5pt solid black'>Quantity</td>                                                                                                                                    \n");
                htmlStr.Append("  <td colspan=3 class=xl8112067>Unit price</td>                                                                                                                                                                         \n");
                htmlStr.Append("  <td colspan=3 class=xl17412067 style='border-right:.5pt solid black'>Amount</td>                                                                                                                                      \n");
                htmlStr.Append(" </tr>                                                                                                                                                                                                                  \n");
                htmlStr.Append(" <tr class=xl8212067 height=20 style='mso-height-source:userset;height:15.0pt'>                                                                                                                                         \n");
                htmlStr.Append("  <td colspan=2 height=20 class=xl17912067 style='height:15.0pt'>A</td>                                                                                                                                                 \n");
                htmlStr.Append("  <td colspan=6 class=xl17912067 style='border-right:.5pt solid black'>B</td>                                                                                                                                           \n");
                htmlStr.Append("  <td class=xl14212067>C</td>                                                                                                                                                                                           \n");
                htmlStr.Append("  <td colspan=3 class=xl17912067 style='border-right:.5pt solid black'>1</td>                                                                                                                                           \n");
                htmlStr.Append("  <td colspan=3 class=xl14212067>2</td>                                                                                                                                                                                 \n");
                htmlStr.Append("  <td colspan=3 class=xl17912067 style='border-right:.5pt solid black'>3 = 1 x                                                                                                                                          \n");
                htmlStr.Append("  2</td>                                                                                                                                                                                                                \n");
                htmlStr.Append(" </tr>																																																					\n");

                v_rowHeight = "26.0pt"; //"26.5pt";
                v_rowHeightEmpty = "22.0pt";
                v_rowHeightNumber = 26;

                v_rowHeightLast = "21.0pt";// "23.5pt";
                v_rowHeightLastNumber = 21;// 26;
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
                        htmlStr.Append(" <tr class=xl8012067 height=25 style='mso-height-source:userset;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>															 \n");
                        htmlStr.Append("  <td colspan=2 height=25 class=xl13312067 width=39 style='border-right:.5pt solid black;                                                    \n");
                        htmlStr.Append("  height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ";width:29pt'>&nbsp;" + dt_d.Rows[v_index][7] + "</td>                                                                                        \n");
                        htmlStr.Append("  <td colspan=6 class=xl13512067 width=289 style='border-right:.5pt solid black;                                                             \n");
                        htmlStr.Append("  border-left:none;width:216pt'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                                                     \n");
                        htmlStr.Append("  <td class=xl11812067 style='border-left:none'>&nbsp;" + dt_d.Rows[v_index][1] + "</td>                                                                     \n");
                        htmlStr.Append("  <td colspan=2 class=xl18112067 style='border-left:none'>&nbsp;" + dt_d.Rows[v_index][2] + "</td>                                                           \n");
                        htmlStr.Append("  <td class=xl11512067>&nbsp;</td>                                                                                                           \n");
                        htmlStr.Append("  <td colspan=2 class=xl13812067 style='border-left:none'>&nbsp;" + dt_d.Rows[v_index][3] + "</td>                                                           \n");
                        htmlStr.Append("  <td class=xl11412067>&nbsp;</td>                                                                                                           \n");
                        htmlStr.Append("  <td colspan=2 class=xl14012067 width=147 style='border-left:none;width:110pt'>&nbsp;" + dt_d.Rows[v_index][4] + "</td>                                     \n");
                        htmlStr.Append("  <td class=xl11412067>&nbsp;</td>                                                                                                           \n");
                        htmlStr.Append(" </tr>                                                                                                                                       \n");


                    }
                    else if (dtR == page[k] - 1)//dong cuoi moi trang
                    {
                        if (k < v_countNumberOfPages - 1) //trang giua
                        {
                            htmlStr.Append("<tr class=xl8012067 height=25 style='mso-height-source:userset;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>																			 \n");
                            htmlStr.Append("	  <td colspan=2 height=25 class=xl13312067 width=39 style='border-right:.5pt solid black;                                                            \n");
                            htmlStr.Append("	  height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ";width:29pt'>&nbsp;" + dt_d.Rows[v_index][7] + "</td>                                                                                                \n");
                            htmlStr.Append("	  <td colspan=6 class=xl13512067 width=289 style='border-right:.5pt solid black;                                                                     \n");
                            htmlStr.Append("	  border-left:none;width:216pt'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                                                             \n");
                            htmlStr.Append("	  <td class=xl11812067 style='border-top:none;border-left:none'>&nbsp;" + dt_d.Rows[v_index][1] + "</td>                                                             \n");
                            htmlStr.Append("	  <td colspan=2 class=xl13812067 style='border-top:none;border-left:none'>&nbsp;" + dt_d.Rows[v_index][2] + "</td>                                                   \n");
                            htmlStr.Append("	  <td class=xl11512067 style='border-top:none'>&nbsp;</td>                                                                                           \n");
                            htmlStr.Append("	  <td colspan=2 class=xl13812067 style='border-left:none'>&nbsp;" + dt_d.Rows[v_index][3] + "</td>                                                                   \n");
                            htmlStr.Append("	  <td class=xl11412067 style='border-top:none'>&nbsp;</td>                                                                                           \n");
                            htmlStr.Append("	  <td colspan=2 class=xl14012067 width=147 style='border-left:none;width:110pt'>&nbsp;" + dt_d.Rows[v_index][4] + "</td>                                             \n");
                            htmlStr.Append("	  <td class=xl11412067 style='border-top:none'>&nbsp;</td>                                                                                           \n");
                            htmlStr.Append(" </tr>                                                                                                                                                   \n");



                        }
                        else // trang cuoi
                        {
                            if (dtR == rowsPerPage - 1) // du 11 dong
                            {
                                htmlStr.Append("<tr class=xl8012067 height=25 style='mso-height-source:userset;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>																			 \n");
                                htmlStr.Append("	  <td colspan=2 height=25 class=xl13312067 width=39 style='border-right:.5pt solid black;                                                            \n");
                                htmlStr.Append("	  height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ";width:29pt'>&nbsp;" + dt_d.Rows[v_index][7] + "</td>                                                                                                \n");
                                htmlStr.Append("	  <td colspan=6 class=xl13512067 width=289 style='border-right:.5pt solid black;                                                                     \n");
                                htmlStr.Append("	  border-left:none;width:216pt'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                                                             \n");
                                htmlStr.Append("	  <td class=xl11812067 style='border-top:none;border-left:none'>&nbsp;" + dt_d.Rows[v_index][1] + "</td>                                                             \n");
                                htmlStr.Append("	  <td colspan=2 class=xl13812067 style='border-top:none;border-left:none'>&nbsp;" + dt_d.Rows[v_index][2] + "</td>                                                   \n");
                                htmlStr.Append("	  <td class=xl11512067 style='border-top:none'>&nbsp;</td>                                                                                           \n");
                                htmlStr.Append("	  <td colspan=2 class=xl13812067 style='border-left:none'>&nbsp;" + dt_d.Rows[v_index][3] + "</td>                                                                   \n");
                                htmlStr.Append("	  <td class=xl11412067 style='border-top:none'>&nbsp;</td>                                                                                           \n");
                                htmlStr.Append("	  <td colspan=2 class=xl14012067 width=147 style='border-left:none;width:110pt'>&nbsp;" + dt_d.Rows[v_index][4] + "</td>                                             \n");
                                htmlStr.Append("	  <td class=xl11412067 style='border-top:none'>&nbsp;</td>                                                                                           \n");
                                htmlStr.Append(" </tr>                                                                                                                                                   \n");
                            }
                            else
                            {
                                htmlStr.Append("<tr class=xl8012067 height=25 style='mso-height-source:userset;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>																			 \n");
                                htmlStr.Append("	  <td colspan=2 height=25 class=xl13312067 width=39 style='border-right:.5pt solid black;                                                            \n");
                                htmlStr.Append("	  height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ";width:29pt'>&nbsp;" + dt_d.Rows[v_index][7] + "</td>                                                                                                \n");
                                htmlStr.Append("	  <td colspan=6 class=xl13512067 width=289 style='border-right:.5pt solid black;                                                                     \n");
                                htmlStr.Append("	  border-left:none;width:216pt'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                                                             \n");
                                htmlStr.Append("	  <td class=xl11812067 style='border-top:none;border-left:none'>&nbsp;" + dt_d.Rows[v_index][1] + "</td>                                                             \n");
                                htmlStr.Append("	  <td colspan=2 class=xl13812067 style='border-top:none;border-left:none'>&nbsp;" + dt_d.Rows[v_index][2] + "</td>                                                   \n");
                                htmlStr.Append("	  <td class=xl11512067 style='border-top:none'>&nbsp;</td>                                                                                           \n");
                                htmlStr.Append("	  <td colspan=2 class=xl13812067 style='border-left:none'>&nbsp;" + dt_d.Rows[v_index][3] + "</td>                                                                   \n");
                                htmlStr.Append("	  <td class=xl11412067 style='border-top:none'>&nbsp;</td>                                                                                           \n");
                                htmlStr.Append("	  <td colspan=2 class=xl14012067 width=147 style='border-left:none;width:110pt'>&nbsp;" + dt_d.Rows[v_index][4] + "</td>                                             \n");
                                htmlStr.Append("	  <td class=xl11412067 style='border-top:none'>&nbsp;</td>                                                                                           \n");
                                htmlStr.Append(" </tr>                                                                                                                                                   \n");
                            }

                        }
                    }
                    else
                    { // dong giua
                      // 
                        htmlStr.Append("<tr class=xl8012067 height=25 style='mso-height-source:userset;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>																			 \n");
                        htmlStr.Append("	  <td colspan=2 height=25 class=xl13312067 width=39 style='border-right:.5pt solid black;                                                            \n");
                        htmlStr.Append("	  height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ";width:29pt'>&nbsp;" + dt_d.Rows[v_index][7] + "</td>                                                                                                \n");
                        htmlStr.Append("	  <td colspan=6 class=xl13512067 width=289 style='border-right:.5pt solid black;                                                                     \n");
                        htmlStr.Append("	  border-left:none;width:216pt'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                                                             \n");
                        htmlStr.Append("	  <td class=xl11812067 style='border-top:none;border-left:none'>&nbsp;" + dt_d.Rows[v_index][1] + "</td>                                                             \n");
                        htmlStr.Append("	  <td colspan=2 class=xl13812067 style='border-top:none;border-left:none'>&nbsp;" + dt_d.Rows[v_index][2] + "</td>                                                   \n");
                        htmlStr.Append("	  <td class=xl11512067 style='border-top:none'>&nbsp;</td>                                                                                           \n");
                        htmlStr.Append("	  <td colspan=2 class=xl13812067 style='border-left:none'>&nbsp;" + dt_d.Rows[v_index][3] + "</td>                                                                   \n");
                        htmlStr.Append("	  <td class=xl11412067 style='border-top:none'>&nbsp;</td>                                                                                           \n");
                        htmlStr.Append("	  <td colspan=2 class=xl14012067 width=147 style='border-left:none;width:110pt'>&nbsp;" + dt_d.Rows[v_index][4] + "</td>                                             \n");
                        htmlStr.Append("	  <td class=xl11412067 style='border-top:none'>&nbsp;</td>                                                                                           \n");
                        htmlStr.Append(" </tr>                                                                                                                                                   \n");

                    }
                    v_index++;
                } //for dtR

                v_spacePerPage = 0;
                if (k < v_countNumberOfPages - 1 && page_index[k] < rowsPerPage)
                {
                    //for (int i = 0; i < rowsPerPage - page_index[k]; i++)
                    //{
                    //v_spacePerPage += v_totalHeightPage;
                    //    v_spacePerPage += v_rowHeightLastNumber;
                    //}
                }
                else if (k < v_countNumberOfPages - 1 && page_index[k] == rowsPerPage)
                {
                    v_spacePerPage = 18;
                }

                if (k == v_countNumberOfPages - 1 && page_index[k] < rowsPerPage) // Trang cuoi khong du dong
                {
                    v_rowHeightEmptyLast = Math.Round(v_totalHeightLastPage / (rowsPerPage - page_index[k]), 2).ToString() + "pt";
                    for (int i = 0; i < rowsPerPage - page_index[k] - 1; i++)
                    {
                        if (i == (rowsPerPage - page_index[k] - 1))
                        {
                            htmlStr.Append("   <tr class=xl8012067 height=25 style='mso-height-source:userset;height:" + v_rowHeightEmptyLast + "'>										  \n");
                            htmlStr.Append("	  <td colspan=2 height=25 class=xl13312067 width=39 style='border-right:.5pt solid black;border-bottom:.5pt solid black;                             \n");
                            htmlStr.Append("	  height:" + v_rowHeightEmptyLast + ";width:29pt;border-bottom:.5pt solid black;border-bottom:.5pt solid black'>&nbsp;</td>                                                                              \n");
                            htmlStr.Append("	  <td colspan=6 class=xl18312067 width=289 style='border-right:.5pt solid black;border-bottom:.5pt solid black;                                      \n");
                            htmlStr.Append("	  border-left:none;width:216pt'></td>                                                                                 \n");
                            htmlStr.Append("	  <td class=xl11812067 style='border-top:none;border-left:none;border-bottom:.5pt solid black'>&nbsp;</td>                                           \n");
                            htmlStr.Append("	  <td colspan=2 class=xl13812067 style='border-top:none;border-left:none;border-bottom:.5pt solid black'>&nbsp;</td>                                 \n");
                            htmlStr.Append("	  <td class=xl11512067 style='border-top:none;border-bottom:.5pt solid black'>&nbsp;</td>                                                            \n");
                            htmlStr.Append("	  <td colspan=2 class=xl13812067 style='border-left:none;border-bottom:.5pt solid black'>&nbsp;</td>                                                 \n");
                            htmlStr.Append("	  <td class=xl11412067 style='border-top:none;border-bottom:.5pt solid black'>&nbsp;</td>                                                            \n");
                            htmlStr.Append("	  <td colspan=2 class=xl14012067 width=147 style='border-left:none;width:110pt;border-bottom:.5pt solid black'>&nbsp;</td>                           \n");
                            htmlStr.Append("	  <td class=xl11412067 style='border-top:none;border-bottom:.5pt solid black'>&nbsp;</td>                                                            \n");
                            htmlStr.Append(" </tr>                                                                                                                    \n");


                        }
                        else
                        {
                            htmlStr.Append("   <tr class=xl8012067 height=25 style='mso-height-source:userset;height:" + v_rowHeightEmptyLast + "'>										  \n");
                            htmlStr.Append("	  <td colspan=2 height=25 class=xl13312067 width=39 style='border-right:.5pt solid black;                             \n");
                            htmlStr.Append("	  height:" + v_rowHeightEmptyLast + ";width:29pt'>&nbsp;</td>                                                                              \n");
                            htmlStr.Append("	  <td colspan=6 class=xl18312067 width=289 style='border-right:.5pt solid black;                                      \n");
                            htmlStr.Append("	  border-left:none;width:216pt'></td>                                                                                 \n");
                            htmlStr.Append("	  <td class=xl11812067 style='border-top:none;border-left:none'>&nbsp;</td>                                           \n");
                            htmlStr.Append("	  <td colspan=2 class=xl13812067 style='border-top:none;border-left:none'>&nbsp;</td>                                 \n");
                            htmlStr.Append("	  <td class=xl11512067 style='border-top:none'>&nbsp;</td>                                                            \n");
                            htmlStr.Append("	  <td colspan=2 class=xl13812067 style='border-left:none'>&nbsp;</td>                                                 \n");
                            htmlStr.Append("	  <td class=xl11412067 style='border-top:none'>&nbsp;</td>                                                            \n");
                            htmlStr.Append("	  <td colspan=2 class=xl14012067 width=147 style='border-left:none;width:110pt'>&nbsp;</td>                           \n");
                            htmlStr.Append("	  <td class=xl11412067 style='border-top:none'>&nbsp;</td>                                                            \n");
                            htmlStr.Append(" </tr>                                                                                                                    \n");
                        }
                    } // for

                }//Trang cuoi 11 dong

                if (k < v_countNumberOfPages - 1)
                {
                    htmlStr.Append("   <tr class=xl7412067 height=25 style='border-bottom: 1pt solid black;mso-height-source:userset;height:" + (v_spacePerPage).ToString() + "pt'>								\n");
                    htmlStr.Append("	  <td colspan=2 height=25 class=xl19112067 width=39 style='border-right:none;                   \n");
                    htmlStr.Append("	  height:" + (v_spacePerPage).ToString() + "pt;width:29pt'>&nbsp;</td>                                                                    \n");
                    htmlStr.Append("	  <td colspan=6 class=xl19312067 width=164 style='border-right:none;border-left:none;width:123pt'>&nbsp;</td>                                  \n");
                    htmlStr.Append("	  <td class=xl11812067 width=78 style='border-right:none;border-left:none;width:55.68pt'>&nbsp;</td>                           \n");
                    htmlStr.Append("	  <td colspan=2 class=xl19412067 style='border-right:none;border-left:none'>&nbsp;</td>                                       \n");
                    htmlStr.Append("	  <td class=xl12312067 width=6 style='border-right:none;border-left:none;width:3.84pt'>&nbsp;</td>                             \n");
                    htmlStr.Append("	  <td colspan=2 class=xl19412067 style='border-right:none;border-left:none'>&nbsp;</td>                                       \n");
                    htmlStr.Append("	  <td class=xl12412067 width=6 style='border-right:none;border-left:none;width:3.84pt'>&nbsp;</td>                             \n");
                    htmlStr.Append("	  <td colspan=2 class=xl19412067 style='border-right:none;border-left:none'>&nbsp;</td>                                       \n");
                    htmlStr.Append("	  <td class=xl11612067 width=6 style='border-left:none;width:3.84pt'>&nbsp;</td>                             \n");
                    htmlStr.Append(" </tr>                                                                                                          \n");

                    htmlStr.Append("	<table  border=0>                                                                                                                                                                                                 \n");
                    htmlStr.Append("		<tr height=5  style='height: 18pt'>                                                                                                                                                                \n");

                    htmlStr.Append("		</tr>      																																														\n");
                    htmlStr.Append("	</table>             																																										\n");

                }


            }// for k                                                                                                                             
            htmlStr.Append(" <tr class=xl7412067 height=26 style='mso-height-source:userset;height:19.89pt'>																		\n");
            htmlStr.Append("  <td height=26 class=xl9812067 width=6 style='height:19.95pt;width:3.84pt'>&nbsp;</td>                                                                 \n");
            htmlStr.Append("  <td class=xl12512067 width=33 style='width:24pt'>&nbsp;</td>                                                                                          \n");
            htmlStr.Append("  <td class=xl12512067 width=70 style='width:49.92pt'>&nbsp;</td>                                                                                       \n");
            htmlStr.Append("  <td class=xl12512067 width=55 style='width:39.36pt'>&nbsp;</td>                                                                                       \n");
            htmlStr.Append("  <td class=xl12512067 width=41 style='width:29.76pt'>&nbsp;</td>                                                                                       \n");
            htmlStr.Append("  <td class=xl12512067 width=41 style='width:29.76pt'>&nbsp;</td>                                                                                       \n");
            htmlStr.Append("  <td class=xl12512067 width=70 style='width:49.92pt'>&nbsp;</td>                                                                                       \n");
            htmlStr.Append("  <td class=xl12512067 width=12 style='width:8.64pt'>&nbsp;</td>                                                                                        \n");
            htmlStr.Append("  <td colspan=7 class=xl18612067 width=265 style='border-right:.5pt solid black;                                                                        \n");
            htmlStr.Append("  width:198pt'>C&#7897;ng ti&#7873;n hàng (<font class='font1012067'>Total                                                                              \n");
            htmlStr.Append("  amount</font><font class='font812067'>) :</font></td>                                                                                                 \n");
            htmlStr.Append("  <td colspan=2 class=xl18812067 style='border-left:none'>&nbsp;" + amount_net + "</td>                                                                     \n");
            htmlStr.Append("  <td class=xl11712067>&nbsp;</td>                                                                                                                      \n");
            htmlStr.Append(" </tr>                                                                                                                                                  \n");
            htmlStr.Append(" <tr class=xl7412067 height=26 style='mso-height-source:userset;height:19.89pt'>                                                                        \n");
            htmlStr.Append("  <td height=26 class=xl9812067 width=6 style='height:19.95pt;border-top:none;                                                                          \n");
            htmlStr.Append("  width:3.84pt'>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("  <td colspan=5 class=xl13212067 width=199 style='width:149pt'><span                                                                                    \n");
            htmlStr.Append("  style='mso-spacerun:yes'> </span>Thu&#7871; su&#7845;t GTGT (<font                                                                                    \n");
            htmlStr.Append("  class='font1012067'>VAT rate</font><font class='font812067'>) :</font>&nbsp;" + dt.Rows[0]["TaxRate"] + "</td>                                                      \n");
            htmlStr.Append("  <td class=xl12712067 width=70 style='border-top:none;width:49.92pt'>&nbsp;</td>                                                                       \n");
            htmlStr.Append("  <td class=xl13212067 width=12 style='border-top:none;width:8.64pt'>&nbsp;</td>                                                                        \n");
            htmlStr.Append("  <td colspan=7 class=xl18612067 width=265 style='border-right:.5pt solid black;                                                                        \n");
            htmlStr.Append("  width:198pt'>Ti&#7873;n thu&#7871; GTGT (<font class='font1012067'>VAT</font><font                                                                    \n");
            htmlStr.Append("  class='font812067'>) :</font></td>                                                                                                                    \n");
            htmlStr.Append("  <td colspan=2 class=xl18812067 style='border-left:none'>&nbsp;" + amount_vat + "</td>                                                                   \n");
            htmlStr.Append("  <td class=xl8412067 style='border-top:none'>&nbsp;</td>                                                                                               \n");
            htmlStr.Append(" </tr>                                                                                                                                                  \n");
            htmlStr.Append(" <tr class=xl7412067 height=26 style='mso-height-source:userset;height:19.89pt'>                                                                        \n");
            htmlStr.Append("  <td height=26 class=xl9812067 width=6 style='height:19.95pt;border-top:none;                                                                          \n");
            htmlStr.Append("  width:3.84pt'>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("  <td class=xl13212067 width=33 style='border-top:none;width:24pt'>&nbsp;</td>                                                                          \n");
            htmlStr.Append("  <td class=xl13212067 width=70 style='border-top:none;width:49.92pt'>&nbsp;</td>                                                                       \n");
            htmlStr.Append("  <td class=xl13212067 width=55 style='border-top:none;width:39.36pt'>&nbsp;</td>                                                                       \n");
            htmlStr.Append("  <td class=xl13212067 width=41 style='border-top:none;width:29.76pt'>&nbsp;</td>                                                                       \n");
            htmlStr.Append("  <td class=xl13212067 width=41 style='border-top:none;width:29.76pt'>&nbsp;</td>                                                                       \n");
            htmlStr.Append("  <td class=xl13212067 width=70 style='border-top:none;width:49.92pt'>&nbsp;</td>                                                                       \n");
            htmlStr.Append("  <td class=xl13212067 width=12 style='border-top:none;width:8.64pt'>&nbsp;</td>                                                                        \n");
            htmlStr.Append("  <td colspan=7 class=xl18612067 width=265 style='border-right:.5pt solid black;                                                                        \n");
            htmlStr.Append("  width:198pt'>T&#7893;ng ti&#7873;n thanh toán (<font class='font1012067'>Total                                                                        \n");
            htmlStr.Append("  payment</font><font class='font812067'>) :</font></td>                                                                                                \n");
            htmlStr.Append("  <td colspan=2 class=xl18812067 style='border-left:none'>&nbsp;" + amount_total + "</td>                                                            \n");
            htmlStr.Append("  <td class=xl8412067>&nbsp;</td>                                                                                                                       \n");
            htmlStr.Append(" </tr>                                                                                                                                                  \n");
            htmlStr.Append(" <tr class=xl7412067 height=24 style='mso-height-source:userset;height:20.73pt'>                                                                        \n");
            htmlStr.Append("  <td height=24 class=xl8312067 style='height:18.0pt;border-top:none'>&nbsp;</td>                                                                       \n");
            htmlStr.Append("  <td colspan=16 class=xl19012067>S&#7889; ti&#7873;n vi&#7871;t b&#7857;ng                                                                             \n");
            htmlStr.Append("  ch&#7919; (<font class='font1012067'>In words</font><font class='font7120671'>):                                                                      \n");
            htmlStr.Append("  &nbsp;" + read_prive + "</font></td>                                                                                                                     \n");
            htmlStr.Append("  <td class=xl8512067 width=6 style='width:3.84pt'>&nbsp;</td>                                                                                          \n");
            htmlStr.Append(" </tr>                                                                                                                                                  \n");
            htmlStr.Append(" <tr height=18 style='mso-height-source:userset;height:16.8pt'>                                                                                        \n");
            htmlStr.Append("  <td height=18 class=xl7912067 style='height:16.8pt'>&nbsp;</td>                                                                                      \n");
            htmlStr.Append("  <td class=xl8612067>&nbsp;</td>                                                                                                                       \n");
            htmlStr.Append("  <td class=xl8612067>&nbsp;</td>                                                                                                                       \n");
            htmlStr.Append("  <td class=xl8612067>&nbsp;</td>                                                                                                                       \n");
            htmlStr.Append("  <td class=xl8712067 width=41 style='width:29.76pt'>&nbsp;</td>                                                                                        \n");
            htmlStr.Append("  <td class=xl8712067 width=41 style='width:29.76pt'>&nbsp;</td>                                                                                        \n");
            htmlStr.Append("  <td class=xl8712067 width=70 style='width:49.92pt'>&nbsp;</td>                                                                                        \n");
            htmlStr.Append("  <td class=xl8712067 width=12 style='width:8.64pt'>&nbsp;</td>                                                                                         \n");
            htmlStr.Append("  <td class=xl8712067 width=78 style='width:55.68pt'>&nbsp;</td>                                                                                        \n");
            htmlStr.Append("  <td class=xl8712067 width=56 style='width:40.32pt'>&nbsp;</td>                                                                                        \n");
            htmlStr.Append("  <td class=xl8712067 width=30 style='width:22.08pt'>&nbsp;</td>                                                                                        \n");
            htmlStr.Append("  <td class=xl8712067 width=6 style='width:3.84pt'>&nbsp;</td>                                                                                          \n");
            htmlStr.Append("  <td class=xl8712067 width=41 style='width:29.76pt'>&nbsp;</td>                                                                                        \n");
            htmlStr.Append("  <td class=xl8712067 width=48 style='width:34.56pt'>&nbsp;</td>                                                                                        \n");
            htmlStr.Append("  <td class=xl8712067 width=6 style='width:3.84pt'>&nbsp;</td>                                                                                          \n");
            htmlStr.Append("  <td class=xl8712067 width=41 style='width:29.76pt'>&nbsp;</td>                                                                                        \n");
            htmlStr.Append("  <td class=xl8712067 width=106 style='width75.84pt'>&nbsp;</td>                                                                                        \n");
            htmlStr.Append("  <td class=xl8812067 width=6 style='width:3.84pt'>&nbsp;</td>                                                                                          \n");
            htmlStr.Append(" </tr>                                                                                                                                                  \n");
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
            htmlStr.Append(" <![if supportMisalignedColumns]>                                                                                                                                   \n");
            htmlStr.Append(" <tr height=0 style='display:none'>                                                                                                                                 \n");
            htmlStr.Append("  <td width=6 style='width:3.84pt'></td>                                                                                                                            \n");
            htmlStr.Append("  <td width=33 style='width:24pt'></td>                                                                                                                             \n");
            htmlStr.Append("  <td width=70 style='width:49.92pt'></td>                                                                                                                          \n");
            htmlStr.Append("  <td width=55 style='width:39.36pt'></td>                                                                                                                          \n");
            htmlStr.Append("  <td width=41 style='width:29.76pt'></td>                                                                                                                          \n");
            htmlStr.Append("  <td width=41 style='width:29.76pt'></td>                                                                                                                          \n");
            htmlStr.Append("  <td width=70 style='width:49.92pt'></td>                                                                                                                          \n");
            htmlStr.Append("  <td width=12 style='width:8.64pt'></td>                                                                                                                           \n");
            htmlStr.Append("  <td width=78 style='width:55.68pt'></td>                                                                                                                          \n");
            htmlStr.Append("  <td width=56 style='width:40.32pt'></td>                                                                                                                          \n");
            htmlStr.Append("  <td width=30 style='width:22.08pt'></td>                                                                                                                          \n");
            htmlStr.Append("  <td width=6 style='width:3.84pt'></td>                                                                                                                            \n");
            htmlStr.Append("  <td width=41 style='width:29.76pt'></td>                                                                                                                          \n");
            htmlStr.Append("  <td width=48 style='width:34.56pt'></td>                                                                                                                          \n");
            htmlStr.Append("  <td width=6 style='width:3.84pt'></td>                                                                                                                            \n");
            htmlStr.Append("  <td width=41 style='width:29.76pt'></td>                                                                                                                          \n");
            htmlStr.Append("  <td width=106 style='width75.84pt'></td>                                                                                                                          \n");
            htmlStr.Append("  <td width=6 style='width:3.84pt'></td>                                                                                                                            \n");
            htmlStr.Append(" </tr>                                                                                                                                                              \n");
            htmlStr.Append(" <![endif]>                                                                                                                                                         \n");
            htmlStr.Append("</table>                                                                                                                                                            \n");
            htmlStr.Append("</body>                                                                                                                                                                     \n");
            htmlStr.Append("</html>                                                                                                                                                                     \n");

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
        private static int countLength(string s)
        {
            int result = 0;
            int max_length = 30;
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