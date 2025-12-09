
using iTextSharp.text.pdf;
using System;
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
    public class WebCash_C
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

            int pos = 5, pos_lv = 20, v_count = 0, count_page = 0, count_page_v = 0, r = 0, x = 0;

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
            }
            else
            {
                lb_amount_trans = "Tổng cộng VND <font class='font45'>(Amount VND):</font>";
                amount_trans = dt.Rows[0]["TOT_AMT_BK_93"].ToString();
                amount_total = dt.Rows[0]["tot_amt_tr_94"].ToString();
                amount_vat = dt.Rows[0]["VAT_TR_AMT_DIS_TR_91"].ToString();
                amount_net = dt.Rows[0]["NET_TR_AMT_DIS_TR_89"].ToString();

                // read_prive = Num2VNText(dt.Rows[0]["amt_tr_read_95"].ToString(), "USD");
            }

            // read_prive = read_prive.Substring(0, 2) + read_prive.Substring(2, read_prive.Length - 2).ToLower();
            if (dt.Rows[0]["taxrate"].ToString().Trim() == "KCT")
            {
                amount_vat = "-";
            }
            /* else
             {
                 amount_vat = dt.Rows[0]["vatamount_display"].ToString();
             }*/

            //read_en = dt.Rows[0]["TotalAmountInWord"].ToString();
            read_prive = dt.Rows[0]["AMOUNT_WORD_VIE"].ToString();//
            int end = 0;
            int count = count_page_v + r;
            double height = 130;
            StringBuilder htmlStr = new StringBuilder("");
            string heigh = "", heigh_d = "";


            htmlStr.Append("<!DOCTYPE html PUBLIC '-//W3C//DTD HTML 4.01 Transitional//EN' 'http://www.w3.org/TR/html4/loose.dtd'>																			 \n");
            htmlStr.Append("<html>                                                                                                                                                                           \n");
            htmlStr.Append("<head>                                                                                                                                                                           \n");
            htmlStr.Append("<meta http-equiv='Content-Type' content='text/html; charset=UTF-8'>                                                                                                              \n");
            htmlStr.Append("<script type='text/javascript' src='${pageContext.request.contextPath}/system/syscommand.js'></script>                                                                           \n");
            htmlStr.Append("<title>Report E-Invoice</title>                                                                                                                                                  \n");
            htmlStr.Append("  <style>@page { size: A4 }</style>                                                                                                                                              \n");
            htmlStr.Append("  <style>                                                                                                                                                                        \n");
            htmlStr.Append("    /*body   { font-family: serif }                                                                                                                                              \n");
            htmlStr.Append("    h1     { font-family: 'Tangerine', cursive; font-size: 40pt; line-height: 18mm}                                                                                              \n");
            htmlStr.Append("    h2, h3 { font-family: 'Tangerine', cursive; font-size: 24pt; line-height: 7mm }                                                                                              \n");
            htmlStr.Append("    h4     { font-size: 13pt; line-height: 1mm }                                                                                                                                 \n");
            htmlStr.Append("    h2 + p { font-size: 18pt; line-height: 7mm }                                                                                                                                 \n");
            htmlStr.Append("    h3 + p { font-size: 14pt; line-height: 7mm }                                                                                                                                 \n");
            htmlStr.Append("    li     { font-size: 11pt; line-height: 5mm }                                                                                                                                 \n");
            htmlStr.Append("                                                                                                                                                                                 \n");
            htmlStr.Append("    h1      { margin: 0 }                                                                                                                                                        \n");
            htmlStr.Append("    h1 + ul { margin: 2mm 0 5mm }                                                                                                                                                \n");
            htmlStr.Append("    h2, h3  { margin: 0 3mm 3mm 0; float: left }                                                                                                                                 \n");
            htmlStr.Append("    h2 + p,                                                                                                                                                                      \n");
            htmlStr.Append("    h3 + p  { margin: 0 0 3mm 50mm }                                                                                                                                             \n");
            htmlStr.Append("    //h4      { margin: 1mm 0 0 2mm; border-bottom: 1px solid black }                                                                                                            \n");
            htmlStr.Append("    h4 + ul { margin: 5mm 0 0 50mm }                                                                                                                                             \n");
            htmlStr.Append("    article { border: 4px double black; padding: 5mm 10mm; border-radius: 3mm }*/                                                                                                \n");
            htmlStr.Append("    body {                                                                                                                                                                       \n");
            htmlStr.Append("       		 color: blue;                                                                                                                                                        \n");
            htmlStr.Append("       		 font-size:100%;                                                                                                                                                     \n");
            htmlStr.Append("       		 background-image: url('assets/Solution.jpg');                                                                                                                       \n");
            htmlStr.Append("		 }                                                                                                                                                                       \n");
            htmlStr.Append("	h1 {                                                                                                                                                                         \n");
            htmlStr.Append("	        color: #00FF00;                                                                                                                                                      \n");
            htmlStr.Append("	}                                                                                                                                                                            \n");
            htmlStr.Append("	p {                                                                                                                                                                          \n");
            htmlStr.Append("	        color: rgb(0,0,255)                                                                                                                                                  \n");
            htmlStr.Append("	}                                                                                                                                                                            \n");
            htmlStr.Append("	                                                                                                                                                                             \n");
            htmlStr.Append("   headline1 {                                                                                                                                                                   \n");
            htmlStr.Append("      background-image: url(assets/Solution.jpg);                                                                                                                                \n");
            htmlStr.Append("      background-repeat: no-repeat;                                                                                                                                              \n");
            htmlStr.Append("      background-position: left top;                                                                                                                                             \n");
            htmlStr.Append("      padding-top:68px;                                                                                                                                                          \n");
            htmlStr.Append("      margin-bottom:50px;                                                                                                                                                        \n");
            htmlStr.Append("   }                                                                                                                                                                             \n");
            htmlStr.Append("   headline2 {                                                                                                                                                                   \n");
            htmlStr.Append("      background-image: url(images/newsletter_headline2.gif);                                                                                                                    \n");
            htmlStr.Append("      background-repeat: no-repeat;                                                                                                                                              \n");
            htmlStr.Append("      background-position: left top;                                                                                                                                             \n");
            htmlStr.Append("      padding-top:68px;                                                                                                                                                          \n");
            htmlStr.Append("   }                                                                                                                                                                             \n");
            htmlStr.Append("   <!--table                                                                                                                                                                     \n");
            htmlStr.Append("	{mso-displayed-decimal-separator:'\\.';                                                                                                                                       \n");
            htmlStr.Append("	mso-displayed-thousand-separator:'\\,';}                                                                                                                                      \n");
            htmlStr.Append("@page                                                                                                                                                                            \n");
            htmlStr.Append("	{margin:.5in .3in .25in .5in;                                                                                                                                                \n");
            htmlStr.Append("	mso-header-margin:.25in;                                                                                                                                                     \n");
            htmlStr.Append("	mso-footer-margin:.25in;}                                                                                                                                                    \n");
            htmlStr.Append(".font6                                                                                                                                                                           \n");
            htmlStr.Append("	{color:#003366;                                                                                                                                                              \n");
            htmlStr.Append("	font-size:11.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                           \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                        \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                         \n");
            htmlStr.Append(".font7                                                                                                                                                                           \n");
            htmlStr.Append("	{color:#003366;                                                                                                                                                              \n");
            htmlStr.Append("	font-size:15.0pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                           \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                        \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                         \n");
            htmlStr.Append(".font8                                                                                                                                                                           \n");
            htmlStr.Append("	{color:#003366;                                                                                                                                                              \n");
            htmlStr.Append("	font-size:9.5pt;                                                                                                                                                             \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                           \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                        \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                         \n");
            htmlStr.Append(".font9                                                                                                                                                                           \n");
            htmlStr.Append("	{color:#003366;                                                                                                                                                              \n");
            htmlStr.Append("	font-size:16.25pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                           \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                        \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                         \n");
            htmlStr.Append(".font14                                                                                                                                                                          \n");
            htmlStr.Append("	{color:#003366;                                                                                                                                                              \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                             \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                           \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                        \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                         \n");
            htmlStr.Append(".font15                                                                                                                                                                          \n");
            htmlStr.Append("	{color:#003366;                                                                                                                                                              \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                             \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                           \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                        \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                         \n");
            htmlStr.Append(".font16                                                                                                                                                                          \n");
            htmlStr.Append("	{color:#003366;                                                                                                                                                              \n");
            htmlStr.Append("	font-size:11.5pt;                                                                                                                                                             \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                           \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                        \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                         \n");
            htmlStr.Append(".font18                                                                                                                                                                          \n");
            htmlStr.Append("	{color:red;                                                                                                                                                                  \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                           \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                        \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                         \n");
            htmlStr.Append(".font19                                                                                                                                                                          \n");
            htmlStr.Append("	{                                                                                                                                                                            \n");
            htmlStr.Append("	                                                                                                                                                                             \n");
            htmlStr.Append("	color:#003366;                                                                                                                                                               \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                             \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                           \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                        \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                         \n");
            htmlStr.Append(".font22                                                                                                                                                                          \n");
            htmlStr.Append("	{                                                                                                                                                                            \n");
            htmlStr.Append("	color:#003366;                                                                                                                                                               \n");
            htmlStr.Append("	font-size:11.5pt;                                                                                                                                                             \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                             \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                           \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                        \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                         \n");
            htmlStr.Append(".font23                                                                                                                                                                          \n");
            htmlStr.Append("	{                                                                                                                                                                            \n");
            htmlStr.Append("	color:#003366;                                                                                                                                                               \n");
            htmlStr.Append("	font-size:11.5pt;                                                                                                                                                             \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                             \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                           \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                        \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                         \n");
            htmlStr.Append(".font28                                                                                                                                                                          \n");
            htmlStr.Append("	{                                                                                                                                                                            \n");
            htmlStr.Append("	color:#003366;                                                                                                                                                               \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                           \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                        \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                         \n");
            htmlStr.Append(".font29                                                                                                                                                                          \n");
            htmlStr.Append("	{color:#003366;                                                                                                                                                              \n");
            htmlStr.Append("	font-size:11.5pt;                                                                                                                                                             \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                             \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                           \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                        \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                         \n");
            htmlStr.Append(".font31                                                                                                                                                                          \n");
            htmlStr.Append("	{color:#003366;                                                                                                                                                              \n");
            htmlStr.Append("	font-size:15.0pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                           \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                        \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                         \n");
            htmlStr.Append(".font32                                                                                                                                                                          \n");
            htmlStr.Append("	{color:purple;                                                                                                                                                               \n");
            htmlStr.Append("	font-size:15.0pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                           \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                        \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                         \n");
            htmlStr.Append(".font33                                                                                                                                                                          \n");
            htmlStr.Append("	{color:red;                                                                                                                                                                  \n");
            htmlStr.Append("	font-size:15.0pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                           \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                        \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                         \n");
            htmlStr.Append(".font45                                                                                                                                                                          \n");
            htmlStr.Append("	{color:#003366;                                                                                                                                                              \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                             \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                           \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                        \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                         \n");
            htmlStr.Append(".font46                                                                                                                                                                          \n");
            htmlStr.Append("	{color:#003366;                                                                                                                                                              \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                             \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                           \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                        \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                         \n");
            htmlStr.Append(".font48                                                                                                                                                                          \n");
            htmlStr.Append("	{color:#003366;                                                                                                                                                              \n");
            htmlStr.Append("	font-size:15.0pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                             \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                           \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                        \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                         \n");
            htmlStr.Append(".font60                                                                                                                                                                          \n");
            htmlStr.Append("	{color:red;                                                                                                                                                                  \n");
            htmlStr.Append("	font-size:11.5pt;                                                                                                                                                             \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                           \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                        \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                         \n");
            htmlStr.Append(".font64                                                                                                                                                                          \n");
            htmlStr.Append("	{color:red;                                                                                                                                                                  \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                           \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                        \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                         \n");
            htmlStr.Append(".font65                                                                                                                                                                          \n");
            htmlStr.Append("	{color:#003366;                                                                                                                                                              \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                           \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                        \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;}	                                                                                                                                                     \n");
            htmlStr.Append("tr                                                                                                                                                                               \n");
            htmlStr.Append("	{mso-height-source:auto;}                                                                                                                                                    \n");
            htmlStr.Append("col                                                                                                                                                                              \n");
            htmlStr.Append("	{mso-width-source:auto;}                                                                                                                                                     \n");
            htmlStr.Append("br                                                                                                                                                                               \n");
            htmlStr.Append("	{mso-data-placement:same-cell;}                                                                                                                                              \n");
            htmlStr.Append(".style43                                                                                                                                                                         \n");
            htmlStr.Append("	{mso-number-format:'_\\(* \\#\\,\\#\\#0\\.00_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\.00\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                                     \n");
            htmlStr.Append("	mso-style-name:Comma;                                                                                                                                                        \n");
            htmlStr.Append("	mso-style-id:3;}                                                                                                                                                             \n");
            htmlStr.Append(".style53                                                                                                                                                                         \n");
            htmlStr.Append("	{color:#0066CC;                                                                                                                                                              \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                             \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                           \n");
            htmlStr.Append("	text-decoration:underline;                                                                                                                                                   \n");
            htmlStr.Append("	text-underline-style:single;                                                                                                                                                 \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	mso-style-name:Hyperlink;                                                                                                                                                    \n");
            htmlStr.Append("	mso-style-id:8;}                                                                                                                                                             \n");
            htmlStr.Append("a:link                                                                                                                                                                           \n");
            htmlStr.Append("	{color:#0066CC;                                                                                                                                                              \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                             \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                           \n");
            htmlStr.Append("	text-decoration:underline;                                                                                                                                                   \n");
            htmlStr.Append("	text-underline-style:single;                                                                                                                                                 \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                         \n");
            htmlStr.Append("a:visited                                                                                                                                                                        \n");
            htmlStr.Append("	{color:purple;                                                                                                                                                               \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                             \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                           \n");
            htmlStr.Append("	text-decoration:underline;                                                                                                                                                   \n");
            htmlStr.Append("	text-underline-style:single;                                                                                                                                                 \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-generic-font-family:auto;                                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                         \n");
            htmlStr.Append(".style0                                                                                                                                                                          \n");
            htmlStr.Append("	{mso-number-format:General;                                                                                                                                                  \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:bottom;                                                                                                                                                       \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                          \n");
            htmlStr.Append("	mso-rotate:0;                                                                                                                                                                \n");
            htmlStr.Append("	mso-background-source:auto;                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern:auto;                                                                                                                                                            \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                            \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                             \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                           \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                        \n");
            htmlStr.Append("	font-family:Arial;                                                                                                                                                           \n");
            htmlStr.Append("	mso-generic-font-family:auto;                                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	border:none;                                                                                                                                                                 \n");
            htmlStr.Append("	mso-protection:locked visible;                                                                                                                                               \n");
            htmlStr.Append("	mso-style-name:Normal;                                                                                                                                                       \n");
            htmlStr.Append("	mso-style-id:0;}                                                                                                                                                             \n");
            htmlStr.Append("td                                                                                                                                                                               \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	padding:0px;                                                                                                                                                                 \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                          \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                            \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                             \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                           \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                        \n");
            htmlStr.Append("	font-family:Arial;                                                                                                                                                           \n");
            htmlStr.Append("	mso-generic-font-family:auto;                                                                                                                                                \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                   \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:bottom;                                                                                                                                                       \n");
            htmlStr.Append("	border:none;                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source:auto;                                                                                                                                                  \n");
            htmlStr.Append("	mso-pattern:auto;                                                                                                                                                            \n");
            htmlStr.Append("	mso-protection:locked visible;                                                                                                                                               \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                          \n");
            htmlStr.Append("	mso-rotate:0;}                                                                                                                                                               \n");
            htmlStr.Append(".xl65                                                                                                                                                                            \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                            \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;}                                                                                                                                          \n");
            htmlStr.Append(".xl66                                                                                                                                                                            \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                            \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                           \n");
            htmlStr.Append(".xl67                                                                                                                                                                            \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:#003366;                                                                                                                                                               \n");
            htmlStr.Append("	font-size:14.25pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                            \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                           \n");
            htmlStr.Append(".xl68                                                                                                                                                                            \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:#003366;                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                            \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                           \n");
            htmlStr.Append(".xl69                                                                                                                                                                            \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                            \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                          \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                           \n");
            htmlStr.Append(".xl70                                                                                                                                                                            \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;}                                                                                                                                                      \n");
            htmlStr.Append(".xl71                                                                                                                                                                            \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;}                                                                                                                                          \n");
            htmlStr.Append(".xl72                                                                                                                                                                            \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;}                                                                                                                                                      \n");
            htmlStr.Append(".xl73                                                                                                                                                                            \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:15.0pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl74                                                                                                                                                                            \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:#003366;                                                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;}                                                                                                                                                      \n");
            htmlStr.Append(".xl75                                                                                                                                                                            \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:#003366;                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;}                                                                                                                                                      \n");
            htmlStr.Append(".xl76                                                                                                                                                                            \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                          \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                           \n");
            htmlStr.Append(".xl77                                                                                                                                                                            \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:17.5pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;}                                                                                                                                                      \n");
            htmlStr.Append(".xl78                                                                                                                                                                            \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	                                                                                                                                                                             \n");
            htmlStr.Append("	color:#003366;                                                                                                                                                               \n");
            htmlStr.Append("	font-size:11.5pt;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;}                                                                                                                                                      \n");
            htmlStr.Append(".xl79                                                                                                                                                                            \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:#003366;                                                                                                                                                               \n");
            htmlStr.Append("	font-size:11.5pt;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;}                                                                                                                                                      \n");
            htmlStr.Append(".xl80                                                                                                                                                                            \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:15.0pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;}                                                                                                                                                      \n");
            htmlStr.Append(".xl81                                                                                                                                                                            \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:1.0pt;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                            \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                           \n");
            htmlStr.Append(".xl82                                                                                                                                                                            \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:purple;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                            \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                            \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                          \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                               \n");
            htmlStr.Append(".xl83                                                                                                                                                                            \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                         \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;}                                                                                                                                          \n");
            htmlStr.Append(".xl84                                                                                                                                                                            \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                         \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                           \n");
            htmlStr.Append(".xl85                                                                                                                                                                            \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:15.0pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                         \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                            \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl86                                                                                                                                                                            \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:purple;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                         \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                           \n");
            htmlStr.Append(".xl87                                                                                                                                                                            \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                         \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                            \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl88                                                                                                                                                                            \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:purple;                                                                                                                                                                \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                         \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                           \n");
            htmlStr.Append(".xl89                                                                                                                                                                            \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                          \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                         \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                           \n");
            htmlStr.Append(".xl90                                                                                                                                                                            \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:15.0pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                            \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                            \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl91                                                                                                                                                                            \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:15.0pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;}                                                                                                                                                      \n");
            htmlStr.Append(".xl92                                                                                                                                                                            \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:16.25pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:top;}                                                                                                                                                         \n");
            htmlStr.Append(".xl93                                                                                                                                                                            \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:16.25pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;}                                                                                                                                                      \n");
            htmlStr.Append(".xl94                                                                                                                                                                            \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;v                                                                                                                                       \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align:middle;}                                                                                                                                                      \n");
            htmlStr.Append(".xl95                                                                                                                                                                            \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:purple;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	vertical-align:middle;}                                                                                                                                                      \n");
            htmlStr.Append(".xl96                                                                                                                                                                            \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl97                                                                                                                                                                            \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:15.0pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                         \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                           \n");
            htmlStr.Append(".xl98                                                                                                                                                                            \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:15.0pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                            \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                         \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;}                                                                                                                                          \n");
            htmlStr.Append(".xl99                                                                                                                                                                            \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:#003366;                                                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                            \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                         \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                            \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl100                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:#003366;                                                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                            \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                         \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;                                                                                                                                           \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl101                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:#003366;                                                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                            \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                          \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                         \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                            \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl102                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:#003366;                                                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border:1.0pt solid windowtext;                                                                                                                                                \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl103                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:15.0pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                            \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                          \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                         \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                           \n");
            htmlStr.Append(".xl104                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:15.0pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	vertical-align:middle;}                                                                                                                                                      \n");
            htmlStr.Append(".xl105                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:#003366;                                                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                            \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                          \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;                                                                                                                                           \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl106                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:15.0pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;}                                                                                                                                          \n");
            htmlStr.Append(".xl107                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                                            \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                          \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl108                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                             \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                          \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                            \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                          \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;                                                                                                                                           \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                          \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                               \n");
            htmlStr.Append(".xl109                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style43;                                                                                                                                                   \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	mso-number-format:'_\\(* \\#\\,\\#\\#0_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                                              \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                          \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                            \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                          \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;                                                                                                                                           \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                          \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                               \n");
            htmlStr.Append(".xl110                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style43;                                                                                                                                                   \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	mso-number-format:'_\\(* \\#\\,\\#\\#0_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                                              \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                          \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                            \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                            \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                          \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                               \n");
            htmlStr.Append(".xl111                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                                                      \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                            \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                          \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                            \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                            \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                          \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                               \n");
            htmlStr.Append(".xl112                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                            \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                          \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                               \n");
            htmlStr.Append(".xl113                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                                                      \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                            \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;                                                                                                                                           \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                          \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                               \n");
            htmlStr.Append(".xl114                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:15.0pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                          \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                           \n");
            htmlStr.Append(".xl115                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:15.0pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	vertical-align:middle;}                                                                                                                                                      \n");
            htmlStr.Append(".xl116                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                             \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                          \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl117                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                             \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                          \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                            \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl118                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                             \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                          \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;                                                                                                                                           \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                          \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                               \n");
            htmlStr.Append(".xl119                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style43;                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	mso-number-format:'_\\(* \\#\\,\\#\\#0_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                                              \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                          \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;                                                                                                                                           \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                          \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                               \n");
            htmlStr.Append(".xl120                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style43;                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	mso-number-format:'_\\(* \\#\\,\\#\\#0_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                                              \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                          \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                               \n");
            htmlStr.Append(".xl121                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                            \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;                                                                                                                                           \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                          \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                               \n");
            htmlStr.Append(".xl122                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl123                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                          \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;                                                                                                                                           \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                          \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                               \n");
            htmlStr.Append(".xl124                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style43;                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	mso-number-format:'_\\(* \\#\\,\\#\\#0_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                                              \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                          \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;                                                                                                                                           \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                          \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                               \n");
            htmlStr.Append(".xl125                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style43;                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	mso-number-format:'_\\(* \\#\\,\\#\\#0_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                                              \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                          \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                               \n");
            htmlStr.Append(".xl126                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                            \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                          \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                               \n");
            htmlStr.Append(".xl127                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                            \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;                                                                                                                                           \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                          \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                               \n");
            htmlStr.Append(".xl128                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl129                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                             \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                          \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;                                                                                                                                           \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                          \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                               \n");
            htmlStr.Append(".xl130                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style43;                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	mso-number-format:'_\\(* \\#\\,\\#\\#0_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                                              \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                          \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;                                                                                                                                           \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                          \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                               \n");
            htmlStr.Append(".xl131                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                            \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                          \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                               \n");
            htmlStr.Append(".xl132                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                            \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;                                                                                                                                           \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                          \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                               \n");
            htmlStr.Append(".xl133                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style43;                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	mso-number-format:'_\\(* \\#\\,\\#\\#0_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                                              \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                          \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;                                                                                                                                           \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                          \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                               \n");
            htmlStr.Append(".xl134                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style43;                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	mso-number-format:'_\\(* \\#\\,\\#\\#0\\.00_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\.00\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                          \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;                                                                                                                                           \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                          \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                               \n");
            htmlStr.Append(".xl135                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style43;                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	mso-number-format:'_\\(* \\#\\,\\#\\#0\\.00_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\.00\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                          \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                               \n");
            htmlStr.Append(".xl136                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	vertical-align:middle;}                                                                                                                                                      \n");
            htmlStr.Append(".xl137                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;}                                                                                                                                          \n");
            htmlStr.Append(".xl138                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:15.0pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                         \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;}                                                                                                                                          \n");
            htmlStr.Append(".xl139                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                         \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                           \n");
            htmlStr.Append(".xl140                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                          \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                         \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;                                                                                                                                           \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl141                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                         \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                            \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl142                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                         \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;                                                                                                                                           \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl143                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:15.0pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                          \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                         \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                           \n");
            htmlStr.Append(".xl144                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:15.0pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;}                                                                                                                                                      \n");
            htmlStr.Append(".xl145                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:15.0pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                            \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;}                                                                                                                                          \n");
            htmlStr.Append(".xl146                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:15.0pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                            \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                           \n");
            htmlStr.Append(".xl147                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:15.0pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                             \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                          \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                            \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                            \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl148                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:15.0pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                            \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                          \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                           \n");
            htmlStr.Append(".xl149                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;}                                                                                                                                          \n");
            htmlStr.Append(".xl150                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:#0066CC;                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;}                                                                                                                                                      \n");
            htmlStr.Append(".xl151                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:#0066CC;                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl152                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:purple;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;}                                                                                                                                                      \n");
            htmlStr.Append(".xl153                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	mso-number-format:'\\@';                                                                                                                                                      \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                            \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                          \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                          \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                               \n");
            htmlStr.Append(".xl154                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:#0066CC;                                                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	mso-number-format:0%;                                                                                                                                                        \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl155                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                            \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                          \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                          \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                               \n");
            htmlStr.Append(".xl156                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:15.0pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;}                                                                                                                                          \n");
            htmlStr.Append(".xl157                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:#0066CC;                                                                                                                                                               \n");
            htmlStr.Append("	font-size:15.0pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl158                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:15.0pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                          \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                           \n");
            htmlStr.Append(".xl159                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:purple;                                                                                                                                                                \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;}                                                                                                                                                      \n");
            htmlStr.Append(".xl160                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:15.0pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl161                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:15.0pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                         \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                            \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl162                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                         \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                            \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl163                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:15.0pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                            \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                            \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl164                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:15.0pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                            \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                            \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl165                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;}                                                                                                                                          \n");
            htmlStr.Append(".xl166                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                          \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                           \n");
            htmlStr.Append(".xl167                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;v                                                                                                                                       \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;}                                                                                                                                                      \n");
            htmlStr.Append(".xl168                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                             \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;                                                                                                                                           \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl169                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:#0066CC;                                                                                                                                                               \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;}                                                                                                                                                      \n");
            htmlStr.Append(".xl170                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:#0066CC;                                                                                                                                                               \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl171                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style53;                                                                                                                                                   \n");
            htmlStr.Append("	color:#003366;                                                                                                                                                               \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                           \n");
            htmlStr.Append("	text-decoration:underline;                                                                                                                                                   \n");
            htmlStr.Append("	text-underline-style:single;                                                                                                                                                 \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;}                                                                                                                                                      \n");
            htmlStr.Append(".xl172                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:#0066CC;                                                                                                                                                               \n");
            htmlStr.Append("	font-size:10.0pt;                                                                                                                                                             \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	vertical-align:middle;}                                                                                                                                                      \n");
            htmlStr.Append(".xl173                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:red;                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:1.0pt solid #0066CC;                                                                                                                                            \n");
            htmlStr.Append("	border-left:1.0pt solid #0066CC;}                                                                                                                                             \n");
            htmlStr.Append(".xl174                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:red;                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:1.0pt solid #0066CC;                                                                                                                                            \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                           \n");
            htmlStr.Append(".xl175                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:red;                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:1.0pt solid #0066CC;                                                                                                                                             \n");
            htmlStr.Append("	border-bottom:1.0pt solid #0066CC;                                                                                                                                            \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                           \n");
            htmlStr.Append(".xl176                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:#003366;                                                                                                                                                               \n");
            htmlStr.Append("	font-size:10.0pt;                                                                                                                                                             \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                         \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                           \n");
            htmlStr.Append(".xl177                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:#003366;                                                                                                                                                               \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl178                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:red;                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:1.0pt solid #0066CC;                                                                                                                                               \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:1.0pt solid #0066CC;                                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl179                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:red;                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:1.0pt solid #0066CC;                                                                                                                                               \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                            \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl180                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:red;                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:1.0pt solid #0066CC;                                                                                                                                               \n");
            htmlStr.Append("	border-right:1.0pt solid #0066CC;                                                                                                                                             \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                            \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl181                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:#003366;                                                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	vertical-align:middle;}                                                                                                                                                      \n");
            htmlStr.Append(".xl182                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:red;                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:11.5pt;                                                                                                                                                             \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:1.0pt solid #0066CC;                                                                                                                                              \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl183                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:red;                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:11.5pt;                                                                                                                                                             \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl184                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:red;                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:11.5pt;                                                                                                                                                             \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                           \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:1.0pt solid #0066CC;                                                                                                                                             \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                            \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl185                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                             \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                          \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl186                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:#003366;                                                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl187                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	                                                                                                                                                                             \n");
            htmlStr.Append("	color:#003366;                                                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl188                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                             \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                         \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;                                                                                                                                           \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl189                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                             \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                         \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                            \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl190                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                             \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                          \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                         \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                            \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl191                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:12.5pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                             \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                          \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                            \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;                                                                                                                                           \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append("	                                                                                                                                                                             \n");
            htmlStr.Append(".xl192                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                             \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                          \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                            \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                            \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl193                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                             \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                          \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                            \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                          \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                            \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl194                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                             \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;                                                                                                                                           \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl195                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                             \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                          \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl196                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                             \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                          \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                          \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                            \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl197                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:#003366;                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	vertical-align:middle;}                                                                                                                                                      \n");
            htmlStr.Append(".xl198                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:purple;                                                                                                                                                                \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;}                                                                                                                                                      \n");
            htmlStr.Append(".xl199                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                             \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                          \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append("                                                                                                                                                                                 \n");
            htmlStr.Append(".xl200                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	                                                                                                                                                                             \n");
            htmlStr.Append("	color:#003366;                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                         \n");
            htmlStr.Append(".xl201                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:#003366;                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                         \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                            \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                          \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                               \n");
            htmlStr.Append(".xl202                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:red;                                                                                                                                                                   \n");
            htmlStr.Append("	font-size:22.5pt;                                                                                                                                                            \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                          \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                               \n");
            htmlStr.Append(".xl203                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	                                                                                                                                                                             \n");
            htmlStr.Append("	color:#003366;                                                                                                                                                               \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                       \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                          \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                               \n");
            htmlStr.Append(".xl204                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:#003366;                                                                                                                                                               \n");
            htmlStr.Append("	font-size:11.5pt;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                           \n");
            htmlStr.Append("	vertical-align:middle;}                                                                                                                                                      \n");
            htmlStr.Append(".xl205                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                    \n");
            htmlStr.Append("	color:#003366;                                                                                                                                                               \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                             \n");
            htmlStr.Append("	font-family:'Times New Roman', serif;                                                                                                                                        \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                          \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                             \n");
            htmlStr.Append("	vertical-align:middle;}                                                                                                                                                      \n");
            htmlStr.Append("-->                                                                                                                                                                              \n");
            htmlStr.Append("</style>                                                                                                                                                                         \n");
            htmlStr.Append("</head>                                                                                                                                                                          \n");
            htmlStr.Append("<body class='A4'>                                                                                                                                                                \n");

            v_index = 0;
            string v_titlePageNumber = "";
            double v_spacePerPage = 0;

            string v_rowHeight = "50pt"; //"26.5pt";
            string v_rowHeightEmpty = "50pt";
            double v_rowHeightNumber = 50;

            string v_rowHeightLast = "50pt";// "23.5pt";
            double v_rowHeightLastNumber = 50;// 23.5;
            string v_rowHeightEmptyLast = "50pt"; //"23.5pt";

            bool vlongItemName = false;

            double v_totalHeightLastPage = 153.5; // 173.5; //203.5 243.5.5;

            double v_totalHeightPage = 570;//   540;

            for (int k = 0; k < v_countNumberOfPages; k++)
            {
                v_totalHeightPage = 530;// 510;  430

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

                htmlStr.Append("<table x:str border=0 cellpadding=0 cellspacing=0 width=713 style='border-collapse:                                                                                              \n");
                htmlStr.Append(" collapse;table-layout:fixed;width:670pt'>                                                                                                                                       \n");
                htmlStr.Append(" <col class=xl70 width=7 style='mso-width-source:userset;mso-width-alt:256;                                                                                                      \n");
                htmlStr.Append(" width:6.25pt'>                                                                                                                                                                     \n");
                htmlStr.Append(" <col class=xl70 width=32 style='mso-width-source:userset;mso-width-alt:1170;                                                                                                    \n");
                htmlStr.Append(" width:30pt'>                                                                                                                                                                    \n");
                htmlStr.Append(" <col class=xl70 width=159 style='mso-width-source:userset;mso-width-alt:5814;                                                                                                   \n");
                htmlStr.Append(" width:148.75pt'>                                                                                                                                                                   \n");
                htmlStr.Append(" <col class=xl70 width=80 style='mso-width-source:userset;mso-width-alt:2925;                                                                                                    \n");
                htmlStr.Append(" width:75pt'>                                                                                                                                                                    \n");
                htmlStr.Append(" <col class=xl70 width=61 style='mso-width-source:userset;mso-width-alt:2230;                                                                                                    \n");
                htmlStr.Append(" width:57.5pt'>                                                                                                                                                                    \n");
                htmlStr.Append(" <col class=xl70 width=54 style='mso-width-source:userset;mso-width-alt:1974;                                                                                                    \n");
                htmlStr.Append(" width:51.25pt'>                                                                                                                                                                    \n");
                htmlStr.Append(" <col class=xl70 width=61 style='mso-width-source:userset;mso-width-alt:2230;                                                                                                    \n");
                htmlStr.Append(" width:57.5pt'>                                                                                                                                                                    \n");
                htmlStr.Append(" <col class=xl70 width=33 style='mso-width-source:userset;mso-width-alt:1206;                                                                                                    \n");
                htmlStr.Append(" width:31.25pt'>                                                                                                                                                                    \n");
                htmlStr.Append(" <col class=xl70 width=82 style='mso-width-source:userset;mso-width-alt:2998;                                                                                                    \n");
                htmlStr.Append(" width:77.5pt'>                                                                                                                                                                    \n");
                htmlStr.Append(" <col class=xl70 width=12 style='mso-width-source:userset;mso-width-alt:438;                                                                                                     \n");
                htmlStr.Append(" width:11.25pt'>                                                                                                                                                                     \n");
                htmlStr.Append(" <col class=xl70 width=124 style='mso-width-source:userset;mso-width-alt:4534;                                                                                                   \n");
                htmlStr.Append(" width:116.25pt'>                                                                                                                                                                    \n");
                htmlStr.Append(" <col class=xl70 width=8 style='mso-width-source:userset;mso-width-alt:292;                                                                                                      \n");
                htmlStr.Append(" width:7.5pt'>                                                                                                                                                                     \n");
                htmlStr.Append(" <tr height=32 style='mso-height-source:userset;height:30.0pt'>                                                                                                                  \n");
                htmlStr.Append("  <td height=32 class=xl65 width=7 style='height:30.0pt;width:6.25pt'>&nbsp;</td>                                                                                                   \n");
                if (dt.Rows[0]["InvoiceSerialNo"].ToString() != "C23TWG")
                {
                    htmlStr.Append("  <td width=32 class=xl66 style='width:30pt;border-top:1.0pt solid windowtext;' align=left valign=top><![if !vml]><span style='mso-ignore:vglayout;                                                                                                            \n");
                    htmlStr.Append("  position:absolute;z-index:1;margin-left:1px;margin-top:14px;width:230px;                                                                                                       \n");
                    htmlStr.Append("  height:108px'><img width=230 height=87 src='D:/webproject/e-invoice-ws/02.Web/EInvoice/img/WEBCASH-LG_C23TWV.png'                                                                             \n");
                    htmlStr.Append("  v:shapes='Picture_x0020_1'></span><![endif]><span style='mso-ignore:vglayout2'>	                                                                                             \n");
                    htmlStr.Append("	<table cellpadding=0 cellspacing=0>                                                                                                                                          \n");
                    htmlStr.Append("   <tr>                                                                                                                                                                          \n");
                    htmlStr.Append("    <td height=32  width=32 style='height:30.0pt;width:30pt'>&nbsp;</td>                                                                                               \n");
                    htmlStr.Append("   </tr>                                                                                                                                                                         \n");
                    htmlStr.Append("  </table>                                                                                                                                                                       \n");
                    htmlStr.Append("  </span></td>                                                                                                                                                                   \n");

                }
                else
                {
                    htmlStr.Append("  <td width=32 class=xl66 style='width:30pt;border-top:1.0pt solid windowtext;' align=left valign=top><![if !vml]><span style='mso-ignore:vglayout;                                                                                                            \n");
                    htmlStr.Append("  position:absolute;z-index:1;margin-left:1px;margin-top:14px;width:230px;                                                                                                       \n");
                    htmlStr.Append("  height:108px'><img width=230 height=87 src='D:/webproject/e-invoice-ws/02.Web/EInvoice/img/WEBCASH-LOGO_001.png'                                                                             \n");
                    htmlStr.Append("  v:shapes='Picture_x0020_1'></span><![endif]><span style='mso-ignore:vglayout2'>	                                                                                             \n");
                    htmlStr.Append("	<table cellpadding=0 cellspacing=0>                                                                                                                                          \n");
                    htmlStr.Append("   <tr>                                                                                                                                                                          \n");
                    htmlStr.Append("    <td height=32  width=32 style='height:30.0pt;width:30pt'>&nbsp;</td>                                                                                               \n");
                    htmlStr.Append("   </tr>                                                                                                                                                                         \n");
                    htmlStr.Append("  </table>                                                                                                                                                                       \n");
                    htmlStr.Append("  </span></td>                                                                                                                                                                   \n");
                }

                htmlStr.Append("  <td class=xl66 width=159 style='width:148.75pt'>&nbsp;</td>                                                                                                                       \n");
                htmlStr.Append("  <td class=xl67 width=80 style='width:75pt'><font class='font6'><b>" + dt.Rows[0]["Seller_Name_112"] + "  </font><font class='font8'> <i>(" + dt.Rows[0]["Seller_LName_113"] + " )</i></b></font></td>                  \n");
                htmlStr.Append("  <td class=xl68 width=61 style='width:57.5pt'>&nbsp;</td>                                                                                                                         \n");
                htmlStr.Append("  <td class=xl68 width=54 style='width:51.25pt'>&nbsp;</td>                                                                                                                         \n");
                htmlStr.Append("  <td class=xl68 width=61 style='width:57.5pt'>&nbsp;</td>                                                                                                                         \n");
                htmlStr.Append("  <td class=xl68 width=33 style='width:31.25pt'>&nbsp;</td>                                                                                                                         \n");
                htmlStr.Append("  <td class=xl68 width=82 style='width:77.5pt'>&nbsp;</td>                                                                                                                         \n");
                htmlStr.Append("  <td class=xl68 width=12 style='width:11.25pt'>&nbsp;</td>                                                                                                                          \n");
                htmlStr.Append("  <td class=xl68 width=124 style='width:116.25pt'>&nbsp;</td>                                                                                                                        \n");
                htmlStr.Append("  <td class=xl69 width=8 style='width:7.5pt'>&nbsp;</td>                                                                                                                           \n");
                htmlStr.Append(" </tr>                                                                                                                                                                           \n");
                htmlStr.Append(" <tr height=18 style='height:17.2pt'>                                                                                                                                           \n");
                htmlStr.Append("  <td height=18 class=xl71 style='height:17.2pt'>&nbsp;</td>                                                                                                                    \n");
                htmlStr.Append("  <td class=xl72></td>                                                                                                                                                           \n");
                htmlStr.Append("  <td class=xl73></td>                                                                                                                                                           \n");
                htmlStr.Append("  <td class=xl74 colspan=5 style='mso-ignore:colspan'><font class='font14'>Mã                                                                                                    \n");
                htmlStr.Append("  s&#7889; thu&#7871; (</font><font class='font15'>Tax code</font><font                                                                                                          \n");
                htmlStr.Append("  class='font14'>) :</font><font class='font65'>" + dt.Rows[0]["Seller_TaxCode"] + " </font></td>                                                                                              \n");
                htmlStr.Append("  <td colspan=3 class=xl74 style='mso-ignore:colspan'></td>                                                                                                                      \n");
                htmlStr.Append("  <td class=xl76>&nbsp;</td>                                                                                                                                                     \n");
                htmlStr.Append(" </tr>                                                                                                                                                                           \n");
                htmlStr.Append(" <tr height=18 style='mso-height-source:userset;height:16.875pt'>                                                                                                                  \n");
                htmlStr.Append("  <td height=18 class=xl71 style='height:16.875pt'>&nbsp;</td>                                                                                                                     \n");
                htmlStr.Append("  <td colspan=2 class=xl73 style='mso-ignore:colspan'></td>                                                                                                                      \n");
                htmlStr.Append("  <td colspan=8 class=xl200 width=507 style='width:382pt'>&#272;&#7883;a                                                                                                         \n");
                htmlStr.Append("  ch&#7881; <i>(Address)</i><font class='font19'> : " + dt.Rows[0]["Seller_Address"] + " </font></td>                                                                                          \n");
                htmlStr.Append("  <td class=xl76>&nbsp;</td>                                                                                                                                                     \n");
                htmlStr.Append(" </tr>                                                                                                                                                                           \n");
                htmlStr.Append(" <tr height=18 style='mso-height-source:userset;height:16.875pt'>                                                                                                                  \n");
                htmlStr.Append("  <td height=18 class=xl71 style='height:16.875pt'>&nbsp;</td>                                                                                                                     \n");
                htmlStr.Append("  <td class=xl77></td>                                                                                                                                                           \n");
                htmlStr.Append("  <td class=xl72></td>                                                                                                                                                           \n");
                htmlStr.Append("  <td class=xl78 colspan=8 dir=LTR style='mso-ignore:colspan'>&#272;i&#7879;n                                                                                                    \n");
                htmlStr.Append("  tho&#7841;i(<font class='font22'>Tel</font><font class='font23'>) : " + dt.Rows[0]["Seller_Tel"] + " - Fax: " + dt.Rows[0]["Seller_Fax"] + " - Website: " + dt.Rows[0]["seller_website"] + "</font></td>                   \n");
                htmlStr.Append("  <td class=xl76>&nbsp;</td>                                                                                                                                                     \n");
                htmlStr.Append(" </tr>                                                                                                                                                                           \n");
                htmlStr.Append(" <tr height=21 style='height:19.6875pt'>                                                                                                                                           \n");
                htmlStr.Append("  <td height=21 class=xl71 style='height:19.6875pt'>&nbsp;</td>                                                                                                                    \n");
                htmlStr.Append("  <td class=xl80></td>                                                                                                                                                           \n");
                htmlStr.Append("  <td class=xl72></td>                                                                                                                                                           \n");
                htmlStr.Append("  <td colspan=8 class=xl201 dir=LTR>Số tài khoản(A/C No) : <font                                                                                                                 \n");
                htmlStr.Append("  class='font16'>" + dt.Rows[0]["Seller_AccountNo"] + " " + dt.Rows[0]["BANK_NM78"] + " </font></td>                                                                                                                             \n");
                htmlStr.Append("  <td class=xl76>&nbsp;</td>                                                                                                                                                     \n");
                htmlStr.Append(" </tr>                                                                                                                                                                           \n");
                htmlStr.Append("                                                                                                                                                                                 \n");
                htmlStr.Append("  <tr height=0.2 style='mso-height-source:userset;height:0.25pt'>                                                                                                                \n");
                htmlStr.Append("  <td height=0.2 class=xl65 style='height:0.2pt'>&nbsp;</td>                                                                                                                     \n");
                htmlStr.Append("  <td class=xl81>&nbsp;</td>                                                                                                                                                     \n");
                htmlStr.Append("  <td class=xl66>&nbsp;  \n");

                htmlStr.Append("  <span style='mso-ignore:vglayout;																				\n");
                htmlStr.Append("    position:absolute;z-index:2;margin-left:4px;margin-top:10px;width:97px;                                         \n");
                htmlStr.Append("    height:93px'>                                                                                                  \n");
                htmlStr.Append("    <img width=97 height=93                                                                                        \n");
                htmlStr.Append("    src='D:/webproject/e-invoice-ws/02.Web/EInvoice/img/DNC_004.png'                                             \n");
                htmlStr.Append("    v:shapes='_x0000_s3228'>                                                                                       \n");
                htmlStr.Append("    </span>                                                                                                        \n");
                htmlStr.Append("    <![endif]>                                                                                                     \n");
                htmlStr.Append("    <span style='mso-ignore:vglayout2'>                                                                            \n");
                htmlStr.Append("    <table cellpadding=0 cellspacing=0>                                                                            \n");
                htmlStr.Append("     <tr>                                                                                                          \n");
                htmlStr.Append("  	<td height=13 class=xl7221177 width=164 style='height:10.05pt;width:110.7pt'>&nbsp;</td>                        \n");
                htmlStr.Append("     </tr>                                                                                                         \n");
                htmlStr.Append("    </table>                                                                                                       \n");
                htmlStr.Append("    </span>   </td>                                                                                                                                                     \n");
                htmlStr.Append("  <td class=xl82 style='border-top:none'>&nbsp;</td>                                                                                                                             \n");
                htmlStr.Append("  <td class=xl82 style='border-top:none'>&nbsp;</td>                                                                                                                             \n");
                htmlStr.Append("  <td class=xl82 style='border-top:none'>&nbsp;</td>                                                                                                                             \n");
                htmlStr.Append("  <td class=xl82 style='border-top:none'>&nbsp;</td>                                                                                                                             \n");
                htmlStr.Append("  <td class=xl82 style='border-top:none'>&nbsp;</td>                                                                                                                             \n");
                htmlStr.Append("  <td class=xl82 style='border-top:none'>&nbsp;</td>                                                                                                                             \n");
                htmlStr.Append("  <td class=xl82 style='border-top:none'>&nbsp;</td>                                                                                                                             \n");
                htmlStr.Append("  <td class=xl82 style='border-top:none'>&nbsp;</td>                                                                                                                             \n");
                htmlStr.Append("  <td class=xl69>&nbsp;</td>                                                                                                                                                     \n");
                htmlStr.Append(" </tr>                                                                                                                                                                           \n");
                htmlStr.Append(" <tr height=31 style='height:29.0625pt'>                                                                                                                                           \n");
                htmlStr.Append("  <td height=31 class=xl71 style='height:29.0625pt'>&nbsp;</td>                                                                                                                    \n");
                htmlStr.Append("  <td class=xl80></td>                                                                                                                                                           \n");
                htmlStr.Append("  <td class=xl72></td>                                                                                                                                                           \n");
                htmlStr.Append("  <td colspan=5 class=xl202>HÓA ĐƠN GIÁ TRỊ GIA TĂNG</td>                                                                                                                        \n");
                htmlStr.Append("  <td class=xl70></td>                                                                                                                                                           \n");
                htmlStr.Append("  <td colspan=2 class=xl203></td>                                                                                            \n");
                htmlStr.Append("  <td class=xl76>&nbsp;</td>                                                                                                                                                     \n");
                htmlStr.Append(" </tr>                                                                                                                                                                           \n");
                htmlStr.Append(" <tr height=24 style='mso-height-source:userset;height:22.5pt'>                                                                                                                  \n");
                htmlStr.Append("  <td height=24 class=xl71 style='height:22.5pt'>&nbsp;</td>                                                                                                                     \n");
                htmlStr.Append("  <td class=xl80></td>                                                                                                                                                           \n");
                htmlStr.Append("  <td class=xl72></td>                                                                                                                                                           \n");
                htmlStr.Append("  <td colspan=5 class=xl204>HÓA ĐƠN CHUYỂN ĐỔI TỪ HÓA ĐƠN ĐIỆN TỬ)</td>                                                                                                                                         \n");
                htmlStr.Append("  <td class=xl70></td>                                                                                                                                                           \n");
                htmlStr.Append("  <td colspan=2 class=xl205><font class='font14'>Ký hiệu</font><font                                                                                                             \n");
                htmlStr.Append("  class='font16'> :" + dt.Rows[0]["templateCode"] + "" + dt.Rows[0]["InvoiceSerialNo"] + " </font></td>                                                                                                                                \n");
                htmlStr.Append("  <td class=xl76>&nbsp;</td>                                                                                                                                                     \n");
                htmlStr.Append(" </tr>                                                                                                                                                                           \n");
                htmlStr.Append(" <tr height=26 style='mso-height-source:userset;height:24.375pt'>                                                                                                                  \n");
                htmlStr.Append("  <td height=26 class=xl71 style='height:24.375pt'>&nbsp;</td>                                                                                                                     \n");
                htmlStr.Append("  <td class=xl72></td>                                                                                                                                                           \n");
                htmlStr.Append("  <td class=xl73></td>                                                                                                                                                           \n");
                htmlStr.Append("  <td colspan=5 class=xl197>Ngày <font class='font16'>" + dt.Rows[0]["invoiceissueddate_dd"] + " </font><font                                                                                                 \n");
                htmlStr.Append("  class='font14'> tháng </font><font class='font16'>" + dt.Rows[0]["invoiceissueddate_mm"] + " </font><font                                                                                                   \n");
                htmlStr.Append("  class='font14'> năm </font><font class='font16'>" + dt.Rows[0]["invoiceissueddate_yyyy"] + " </font></td>                                                                                                   \n");
                htmlStr.Append("  <td class=xl70></td>                                                                                                                                                           \n");
                htmlStr.Append("  <td colspan=2 class=xl198><font class='font14'>Số </font><font                                                                                                                 \n");
                htmlStr.Append("  class='font16'>:</font><font class='font31'><span style='mso-spacerun:yes'>                                                                                                    \n");
                htmlStr.Append("  </span></font><font class='font32'><span style='mso-spacerun:yes'>   </span></font><font                                                                                       \n");
                htmlStr.Append("  class='font33'>" + dt.Rows[0]["InvoiceNumber"] + " </font></td>                                                                                                                                  \n");
                htmlStr.Append("  <td class=xl76>&nbsp;</td>                                                                                                                                                     \n");
                htmlStr.Append(" </tr>                                                                                                                                                                           \n");
                htmlStr.Append("                                                                                                                                                                                 \n");
                htmlStr.Append("                                                                                                                                                                                 \n");
                htmlStr.Append(" <tr height=25 style='mso-height-source:userset;height:23.6875pt'>                                                                                                                 \n");
                htmlStr.Append("  <td height=25 class=xl71 style='height:23.6875pt;border-top:1.0pt solid black;'>&nbsp;</td>                                                                                       \n");
                htmlStr.Append("  <td class=xl75  style='border-top:1.0pt solid black;'>Họ tên người mua hàng <i>(Customer's name)</i> :</font></td>                                                              \n");
                htmlStr.Append("  <td colspan=2 class=xl91 style='mso-ignore:colspan;border-top:1.0pt solid black;'></td>                                                                                         \n");
                htmlStr.Append("  <td colspan=2 class=xl92 style='mso-ignore:colspan;border-top:1.0pt solid black;'>" + dt.Rows[0]["Buyer"] + "</td>                                                                       \n");
                htmlStr.Append("  <td class=xl93style='mso-ignore:colspan;border-top:1.0pt solid black;'></td>                                                                                                    \n");
                htmlStr.Append("  <td colspan=4 class=xl72 style='mso-ignore:colspan;border-top:1.0pt solid black;'></td>                                                                                         \n");
                htmlStr.Append("  <td class=xl76 style='mso-ignore:colspan;border-top:1.0pt solid black;'>&nbsp;</td>                                                                                             \n");
                htmlStr.Append(" </tr>                                                                                                                                                                           \n");
                htmlStr.Append(" <tr height=21 style='mso-height-source:userset;height:15.95pt'>                                                                                                                 \n");
                htmlStr.Append("  <td height=21 class=xl71 style='height:15.95pt'>&nbsp;</td>                                                                                                                    \n");
                htmlStr.Append("  <td class=xl75>Tên đơn vị <i>(Company's name)</i> :</td>                                                                                                                       \n");
                htmlStr.Append("  <td class=xl72></td>                                                                                                                                                           \n");
                htmlStr.Append("  <td colspan=8 class=xl199>" + dt.Rows[0]["BuyerLegalName"] + " </td>                                                                                                                              \n");
                htmlStr.Append("  <td class=xl76>&nbsp;</td>                                                                                                                                                     \n");
                htmlStr.Append(" </tr>                                                                                                                                                                           \n");
                htmlStr.Append(" <tr height=24 style='mso-height-source:userset;height:22.5pt'>                                                                                                                  \n");
                htmlStr.Append("  <td height=24 class=xl71 style='height:22.5pt'>&nbsp;</td>                                                                                                                     \n");
                htmlStr.Append("  <td class=xl75>Địa chỉ <i>( Address)</i> :</td>                                                                                                                                \n");
                htmlStr.Append("  <td class=xl72></td>                                                                                                                                                           \n");
                htmlStr.Append("  <td colspan=8  class=xl199>" + dt.Rows[0]["BuyerAddress"] + " </td>                                                                                                                                \n");
                htmlStr.Append("  <td class=xl76>&nbsp;</td>                                                                                                                                                     \n");
                htmlStr.Append(" </tr>                                                                                                                                                                           \n");
                htmlStr.Append("                                                                                                                                                                                 \n");
                htmlStr.Append("<!-- tr height=24 style='mso-height-source:userset;height:22.5pt'>                                                                                                               \n");
                htmlStr.Append("  <td height=24 class=xl71 style='height:22.5pt;'>&nbsp;</td>                                                                                                                    \n");
                htmlStr.Append("  <td colspan=2 class=xl121 style='mso-ignore:colspan;border-left:none;border-right:none;'></td>                                                                                 \n");
                htmlStr.Append(" <td class=xl76>&nbsp;</td>                                                                                                                                                      \n");
                htmlStr.Append(" </tr-->                                                                                                                                                                         \n");
                htmlStr.Append(" <tr height=24 style='mso-height-source:userset;height:22.5pt'>                                                                                                                  \n");
                htmlStr.Append("  <td height=24 class=xl71 style='height:22.5pt'>&nbsp;</td>                                                                                                                     \n");
                htmlStr.Append("  <td class=xl75>Số tài khoản <i>(Account no)</i> :</td>                                                                                                                         \n");
                htmlStr.Append("  <td colspan=2 class=xl72 style='mso-ignore:colspan'></td>                                                                                                                      \n");
                htmlStr.Append("  <td colspan=2 class=xl94 style='mso-ignore:colspan'>" + dt.Rows[0]["buyeraccountno"] + " </td>                                                                                                    \n");
                htmlStr.Append("  <td colspan=5 class=xl72 style='mso-ignore:colspan'></td>                                                                                                                      \n");
                htmlStr.Append("  <td class=xl76>&nbsp;</td>                                                                                                                                                     \n");
                htmlStr.Append(" </tr>                                                                                                                                                                           \n");

                htmlStr.Append("   <tr height=24 style='mso-height-source:userset;height:18.0pt'>													 \n");
                htmlStr.Append("    <td height=24 class=xl71 style='height:18.0pt'>&nbsp;</td>                                                       \n");
                htmlStr.Append("    <td class=xl75>Hình thức thanh toán <i>(Method of payment)</i> :</td>                                            \n");
                htmlStr.Append("    <td colspan=2 class=xl72 style='mso-ignore:colspan'></td>                                                        \n");
                htmlStr.Append("    <td colspan=2 class=xl199 style='mso-ignore:colspan'>" + dt.Rows[0]["PaymentMethodCK"] + "</td>                  \n");
                htmlStr.Append("    <td colspan=3 class=xl75 style='mso-ignore:colspan'>Mã số thuế <i>(Tax code)</i> :</td>                          \n");
                htmlStr.Append("    <td colspan=2 class=xl96 style='mso-ignore:colspan'>" + dt.Rows[0]["BuyerTaxCode"] + "</td>                      \n");
                htmlStr.Append("    <td class=xl76>&nbsp;</td>                                                                                       \n");
                htmlStr.Append("   </tr>                                                                                                             \n");
                htmlStr.Append("   <tr height=24 style='mso-height-source:userset;height:18.0pt'>                                                    \n");
                htmlStr.Append("    <td height=24 class=xl71 style='height:18.0pt'>&nbsp;</td>                                                       \n");
                htmlStr.Append("    <td class=xl75>Đơn vị tiền tệ <i>(Currency)</i> :</td>                                                           \n");
                htmlStr.Append("    <td colspan=2 class=xl72 style='mso-ignore:colspan'></td>                                                        \n");
                htmlStr.Append("    <td colspan=2 class=xl94>" + dt.Rows[0]["CurrencyCodeUSD"] + "</td>                                              \n");
                htmlStr.Append("    <td colspan=3 class=xl197 style='text-align: left' >Tỷ giá <i>(Exchange Rate)</i> :</td>                         \n");
                htmlStr.Append("    <td class=xl95></td>                                                                                             \n");
                htmlStr.Append("    <td class=xl96 width=124 style='width:93pt'>" + dt.Rows[0]["tr_rate_88"] + "</td>                                \n");
                htmlStr.Append("    <td class=xl76>&nbsp;</td>                                                                                       \n");
                htmlStr.Append("   </tr>                                                                                                             \n");

                htmlStr.Append("                                                                                                                                                                                 \n");
                htmlStr.Append(" <tr class=xl104 height=51 style='height:47.815pt'>                                                                                                                               \n");
                htmlStr.Append("  <td height=51 class=xl98 style='height:47.815pt;'>&nbsp;</td>                                                                                                                   \n");
                htmlStr.Append("  <td class=xl99 width=32 style='width:30pt'>STT<br>                                                                                                                             \n");
                htmlStr.Append("    <font class='font15'>No.</font></td>                                                                                                                                         \n");
                htmlStr.Append("  <td colspan=3 class=xl100 width=300 style='border-right:1.0pt solid black;                                                                                                      \n");
                htmlStr.Append("  width:225pt'>Tên hàng hóa, d&#7883;ch v&#7909;<br>                                                                                                                             \n");
                htmlStr.Append("    <font class='font15'>Description</font></td>                                                                                                                                 \n");
                htmlStr.Append("  <td class=xl102 width=54 style='border-left:none;width:51.25pt'>&#272;&#417;n                                                                                                     \n");
                htmlStr.Append("  v&#7883; tính<br>                                                                                                                                                              \n");
                htmlStr.Append("    <font class='font15'>Unit</font></td>                                                                                                                                        \n");
                htmlStr.Append("  <td class=xl102 width=61 style='border-left:none;width:57.5pt'>S&#7889;                                                                                                          \n");
                htmlStr.Append("  l&#432;&#7907;ng<br>                                                                                                                                                           \n");
                htmlStr.Append("    <font class='font15'>Quantity</font></td>                                                                                                                                    \n");
                htmlStr.Append("  <td colspan=2 class=xl99 width=115 style='width:87pt'>&#272;&#417;n giá<br>                                                                                                    \n");
                htmlStr.Append("    <font class='font15'>Unit price</font></td>                                                                                                                                  \n");
                htmlStr.Append("  <td class=xl99 width=12 style='width:11.25pt'>&nbsp;</td>                                                                                                                          \n");
                htmlStr.Append("  <td class=xl100 width=124 style='width:116.25pt'>Thành                                                                                                                             \n");
                htmlStr.Append("  ti&#7873;n<br>                                                                                                                                                                 \n");
                htmlStr.Append("    <font class='font15'>Amount</font></td>                                                                                                                                      \n");
                htmlStr.Append("  <td class=xl103 >&nbsp;</td>                                                                                                                                                   \n");
                htmlStr.Append(" </tr>                                                                                                                                                                           \n");
                htmlStr.Append(" <tr class=xl104 height=20 style='height:18.75pt'>                                                                                                                                \n");
                htmlStr.Append("  <td height=20 class=xl98 style='height:18.75pt;border-top:none'>&nbsp;</td>                                                                                                     \n");
                htmlStr.Append("  <td class=xl101 width=32 style='border-top:none;width:30pt' x:num>1</td>                                                                                                       \n");
                htmlStr.Append("  <td colspan=3 class=xl100 width=300 style='border-right:1.0pt solid black;                                                                                                      \n");
                htmlStr.Append("  border-left:none;width:225pt' x:num>2</td>                                                                                                                                     \n");
                htmlStr.Append("  <td class=xl105 width=54 style='border-top:none;border-left:none;border-bottom:1.0pt solid black;width:51.25pt'                                                                    \n");
                htmlStr.Append("  x:num>3</td>                                                                                                                                                                   \n");
                htmlStr.Append("  <td class=xl105 width=61 style='border-top:none;border-left:none;border-bottom:1.0pt solid black;width:57.5pt'                                                                    \n");
                htmlStr.Append("  x:num>4</td>                                                                                                                                                                   \n");
                htmlStr.Append("  <td colspan=3 class=xl100 width=127 style='border-right:1.0pt solid black;                                                                                                      \n");
                htmlStr.Append("  border-left:none;width:96pt' x:num>5</td>                                                                                                                                      \n");
                htmlStr.Append("  <td class=xl100 width=124 style='border-top:none;border-left:none;width:116.25pt'>6                                                                                                \n");
                htmlStr.Append("  = 4 x 5</td>                                                                                                                                                                   \n");
                htmlStr.Append("  <td class=xl103 style='border-top:none'>&nbsp;</td>                                                                                                                            \n");
                htmlStr.Append(" </tr>                                                                                                                                                                           \n");
                htmlStr.Append(" <tr>                                                                                                                                                                            \n");
                htmlStr.Append(" <td>                                                                                                                                                                            \n");

                if (dt.Rows[0]["InvoiceSerialNo"].ToString() != "C23TWG")
                {
                    htmlStr.Append(" <!--[if gte vml 1]><v:rect id='Rectangle_x0020_1'                                                                                                                               \n");
                    htmlStr.Append("   o:spid='_x0000_s1030' style='position:absolute;margin-left:.75pt;                                                                                                             \n");
                    htmlStr.Append("   margin-top:.75pt;width:498pt;height:249.75pt;z-index:2;visibility:visible'                                                                                                    \n");
                    htmlStr.Append("   strokecolor='#41719c' strokeweight='1pt' o:insetmode='auto'>                                                                                                                  \n");
                    htmlStr.Append("   <v:fill src='${pageContext.request.contextPath}/assets/solutionbg1.png' o:title='' opacity='11796f'                                                                           \n");
                    htmlStr.Append("    recolor='t' rotate='t' type='frame'/>                                                                                                                                        \n");
                    htmlStr.Append("  </v:rect><![endif]--><![if !vml]><span style='mso-ignore:vglayout;position:                                                                                                    \n");
                    htmlStr.Append("  absolute;z-index:2;margin-left:0px;margin-top:0px;width:891px;height:357px'><img                                                                                               \n");
                    htmlStr.Append("  width=891 height=357 src='D:/webproject/e-invoice-ws/02.Web/EInvoice/img/WEBCASH-LG_C23TWV_BG.png' v:shapes='Rectangle_x0020_1'></span><![endif]>                                            \n");
                    htmlStr.Append("  </td>                                                                                                                                                                          \n");
                    htmlStr.Append(" </tr>                                                                                                                                                                           \n");


                }
                else
                {
                    htmlStr.Append(" <!--[if gte vml 1]><v:rect id='Rectangle_x0020_1'                                                                                                                               \n");
                    htmlStr.Append("   o:spid='_x0000_s1030' style='position:absolute;margin-left:.75pt;                                                                                                             \n");
                    htmlStr.Append("   margin-top:.75pt;width:498pt;height:249.75pt;z-index:2;visibility:visible'                                                                                                    \n");
                    htmlStr.Append("   strokecolor='#41719c' strokeweight='1pt' o:insetmode='auto'>                                                                                                                  \n");
                    htmlStr.Append("   <v:fill src='${pageContext.request.contextPath}/assets/solutionbg1.png' o:title='' opacity='11796f'                                                                           \n");
                    htmlStr.Append("    recolor='t' rotate='t' type='frame'/>                                                                                                                                        \n");
                    htmlStr.Append("  </v:rect><![endif]--><![if !vml]><span style='mso-ignore:vglayout;position:                                                                                                    \n");
                    htmlStr.Append("  absolute;z-index:2;margin-left:0px;margin-top:0px;width:891px;height:357px'><img                                                                                               \n");
                    htmlStr.Append("  width=891 height=357 src='D:/webproject/e-invoice-ws/02.Web/EInvoice/img/WEBCASH-LOGO_002.png' v:shapes='Rectangle_x0020_1'></span><![endif]>                                            \n");
                    htmlStr.Append("  </td>                                                                                                                                                                          \n");
                    htmlStr.Append(" </tr>                                                                                                                                                                           \n");

                }



                v_rowHeight = "50.0pt"; //"26.5pt";
                v_rowHeightEmpty = "50.0pt";
                v_rowHeightNumber = 50;

                v_rowHeightLast = "50.0pt";// "23.5pt";
                v_rowHeightLastNumber = 50;// 23.5;
                v_rowHeightEmptyLast = "50.pt"; //"23.5pt";


                for (int dtR = 0; dtR < page[k]; dtR++)
                {
                    /*if (!vlongItemName && dt_d.Rows[v_index]["ITEM_NAME"].ToString().Length >= 92)
                    {
                        v_rowHeight = "26.0pt"; //"26.5pt";    
                        v_rowHeightLast = "21.0pt"; //"27.5pt";
                        v_rowHeightLastNumber = 26;//27.5;
                        v_rowHeightEmptyLast = "22.0pt"; //"23.0pt";
                        vlongItemName = true;
                    }*/
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

                        htmlStr.Append(" 			 <tr class=xl115 height=53 style='mso-height-source:userset;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                                         \n");
                        htmlStr.Append("			  <td height=53 class=xl106 style='height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ";border-right:none;border-top:none'>&nbsp;</td>                                                                     \n");
                        htmlStr.Append("			  <td class=xl107 width=32 style='width:30pt' x:num>" + dt_d.Rows[v_index][7] + "</td>                                                                                               \n");
                        htmlStr.Append("			  <td colspan=3  class=xl191 width=300 style='border-top:none;border-right:1.0pt solid black;                                                                         \n");
                        htmlStr.Append("			  width:225pt'>&nbsp;" + dt_d.Rows[v_index][0] + "</td>                                                                                                                                    \n");
                        htmlStr.Append("			  <td class=xl108 style='border-top:none;border-left:none'>&nbsp;" + dt_d.Rows[v_index][1] + "</td>                                                                                  \n");
                        htmlStr.Append("			  <td class=xl109 style='border-top:none;border-left:none' >" + dt_d.Rows[v_index][2] + " </td>    \n");
                        htmlStr.Append("			  <td class=xl110 style='border-top:none;border-top:none' >&nbsp;</td>   \n");

                        htmlStr.Append("			  <td class=xl111 style='border-top:none;border-top:none;'>" + dt_d.Rows[v_index][3] + "</td>                                                                                        \n");
                        htmlStr.Append("			  <td class=xl112></td>                                                                                                                                              \n");
                        htmlStr.Append("			  <td class=xl113>" + dt_d.Rows[v_index][4] + "</td>                                                                                                                                 \n");
                        htmlStr.Append("			  <td class=xl114>&nbsp;</td>                                                                                                                                        \n");
                        htmlStr.Append(" 			</tr>                                                                                                                                                                \n");


                    }
                    else if (v_index == page[k] - 1)//dong cuoi moi trang
                    {
                        if (k < v_countNumberOfPages - 1) //trang giua
                        {

                            htmlStr.Append(" 			 <tr class=xl115 height=53 style='mso-height-source:userset;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                                         \n");
                            htmlStr.Append("			  <td height=53 class=xl106 style='height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ";border-right:none;border-top:none'>&nbsp;</td>                                                                     \n");
                            htmlStr.Append("			  <td class=xl107 width=32 style='width:30pt' x:num>" + dt_d.Rows[v_index][7] + "</td>                                                                                               \n");
                            htmlStr.Append("			  <td colspan=3  class=xl191 width=300 style='border-top:none;border-right:1.0pt solid black;                                                                         \n");
                            htmlStr.Append("			  width:225pt'>&nbsp;" + dt_d.Rows[v_index][0] + "</td>                                                                                                                                    \n");
                            htmlStr.Append("			  <td class=xl108 style='border-top:none;border-left:none'>&nbsp;" + dt_d.Rows[v_index][1] + "</td>                                                                                  \n");
                            htmlStr.Append("			  <td class=xl109 style='border-top:none;border-left:none' >" + dt_d.Rows[v_index][2] + " </td>    \n");
                            htmlStr.Append("			  <td class=xl110 style='border-top:none;border-top:none' >&nbsp;</td>   \n");
                            htmlStr.Append("			  <td class=xl111 style='border-top:none;border-top:none;'>" + dt_d.Rows[v_index][3] + "</td>                                                                                        \n");
                            htmlStr.Append("			  <td class=xl112></td>                                                                                                                                              \n");
                            htmlStr.Append("			  <td class=xl113>" + dt_d.Rows[v_index][4] + "</td>                                                                                                                                 \n");
                            htmlStr.Append("			  <td class=xl114>&nbsp;</td>                                                                                                                                        \n");
                            htmlStr.Append(" 			</tr>                                                                                                                                                                \n");

                        }
                        else // trang cuoi
                        {
                            if (v_index == rowsPerPage - 1) // du 11 dong
                            {

                                htmlStr.Append(" 			 <tr class=xl115 height=53 style='mso-height-source:userset;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                                         \n");
                                htmlStr.Append("			  <td height=53 class=xl106 style='height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ";border-right:none;border-top:none'>&nbsp;</td>                                                                     \n");
                                htmlStr.Append("			  <td class=xl107 width=32 style='width:30pt' x:num>" + dt_d.Rows[v_index][7] + "</td>                                                                                               \n");
                                htmlStr.Append("			  <td colspan=3  class=xl191 width=300 style='border-top:none;border-right:1.0pt solid black;                                                                         \n");
                                htmlStr.Append("			  width:225pt'>&nbsp;" + dt_d.Rows[v_index][0] + "</td>                                                                                                                                    \n");
                                htmlStr.Append("			  <td class=xl108 style='border-top:none;border-left:none'>&nbsp;" + dt_d.Rows[v_index][1] + "</td>                                                                                  \n");
                                htmlStr.Append("			  <td class=xl109 style='border-top:none;border-left:none' >" + dt_d.Rows[v_index][2] + " </td>    \n");
                                htmlStr.Append("			  <td class=xl110 style='border-top:none;border-top:none' >&nbsp;</td>   \n");

                                htmlStr.Append("			  <td class=xl111 style='border-top:none;border-top:none;'>" + dt_d.Rows[v_index][3] + "</td>                                                                                        \n");
                                htmlStr.Append("			  <td class=xl112></td>                                                                                                                                              \n");
                                htmlStr.Append("			  <td class=xl113>" + dt_d.Rows[v_index][4] + "</td>                                                                                                                                 \n");
                                htmlStr.Append("			  <td class=xl114>&nbsp;</td>                                                                                                                                        \n");
                                htmlStr.Append(" 			</tr>                                                                                                                                                                \n");


                            }
                            else
                            {

                                htmlStr.Append(" 			 <tr class=xl115 height=53 style='mso-height-source:userset;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                                         \n");
                                htmlStr.Append("			  <td height=53 class=xl106 style='height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ";border-right:none;border-top:none'>&nbsp;</td>                                                                     \n");
                                htmlStr.Append("			  <td class=xl107 width=32 style='width:30pt' x:num>" + dt_d.Rows[v_index][7] + "</td>                                                                                               \n");
                                htmlStr.Append("			  <td colspan=3  class=xl191 width=300 style='border-top:none;border-right:1.0pt solid black;                                                                         \n");
                                htmlStr.Append("			  width:225pt'>&nbsp;" + dt_d.Rows[v_index][0] + "</td>                                                                                                                                    \n");
                                htmlStr.Append("			  <td class=xl108 style='border-top:none;border-left:none'>&nbsp;" + dt_d.Rows[v_index][1] + "</td>                                                                                  \n");
                                htmlStr.Append("			  <td class=xl109 style='border-top:none;border-left:none' >" + dt_d.Rows[v_index][2] + " </td>    \n");
                                htmlStr.Append("			  <td class=xl110 style='border-top:none;border-top:none' >&nbsp;</td>   \n");

                                htmlStr.Append("			  <td class=xl111 style='border-top:none;border-top:none;'>" + dt_d.Rows[v_index][3] + "</td>                                                                                        \n");
                                htmlStr.Append("			  <td class=xl112></td>                                                                                                                                              \n");
                                htmlStr.Append("			  <td class=xl113>" + dt_d.Rows[v_index][4] + "</td>                                                                                                                                 \n");
                                htmlStr.Append("			  <td class=xl114>&nbsp;</td>                                                                                                                                        \n");
                                htmlStr.Append(" 			</tr>                                                                                                                                                                \n");

                            }

                        }
                    }
                    else
                    { // dong giua                                                                                                                                    

                        htmlStr.Append(" 			 <tr class=xl115 height=53 style='mso-height-source:userset;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                                         \n");
                        htmlStr.Append("			  <td height=53 class=xl106 style='height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ";border-right:none;border-top:none'>&nbsp;</td>                                                                     \n");
                        htmlStr.Append("			  <td class=xl107 width=32 style='width:30pt' x:num>" + dt_d.Rows[v_index][7] + "</td>                                                                                               \n");
                        htmlStr.Append("			  <td colspan=3  class=xl191 width=300 style='border-top:none;border-right:1.0pt solid black;                                                                         \n");
                        htmlStr.Append("			  width:225pt'>&nbsp;" + dt_d.Rows[v_index][0] + "</td>                                                                                                                                    \n");
                        htmlStr.Append("			  <td class=xl108 style='border-top:none;border-left:none'>&nbsp;" + dt_d.Rows[v_index][1] + "</td>                                                                                  \n");
                        htmlStr.Append("			  <td class=xl109 style='border-top:none;border-left:none' >" + dt_d.Rows[v_index][2] + " </td>    \n");
                        htmlStr.Append("			  <td class=xl110 style='border-top:none;border-top:none' >&nbsp;</td>   \n");

                        htmlStr.Append("			  <td class=xl111 style='border-top:none;border-top:none;'>" + dt_d.Rows[v_index][3] + "</td>                                                                                        \n");
                        htmlStr.Append("			  <td class=xl112></td>                                                                                                                                              \n");
                        htmlStr.Append("			  <td class=xl113>" + dt_d.Rows[v_index][4] + "</td>                                                                                                                                 \n");
                        htmlStr.Append("			  <td class=xl114>&nbsp;</td>                                                                                                                                        \n");
                        htmlStr.Append(" 			</tr>                                                                                                                                                                \n");

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

                            htmlStr.Append("			 <tr class=xl115 height=53 style='mso-height-source:userset;height:" + v_rowHeightEmptyLast + " ‬'>                                                                                         \n");
                            htmlStr.Append("				  <td height=53 class=xl106 style='height:" + v_rowHeightEmptyLast + "‬;border-right:none;border-top:none'>&nbsp;</td>                                                                 \n");
                            htmlStr.Append("				  <td class=xl107 width=32 style='width:30pt;border-top:none' x:num></td>                                                                                        \n");
                            htmlStr.Append("				  <td colspan=3  class=xl191 width=300 style='border-top:none;border-right:1.0pt solid black;                                                                     \n");
                            htmlStr.Append("				  width:225pt'></td>                                                                                                                                             \n");
                            htmlStr.Append("				  <td class=xl108 style='border-top:none'></td>                                                                                                                  \n");
                            htmlStr.Append("				  <td class=xl109 style='border-top:none' x:num></td>                                                                                                            \n");
                            htmlStr.Append("				  <td class=xl110 style='border-top:none' ></td>                                                                                                                 \n");
                            htmlStr.Append("				  <td class=xl111 style='border-top:none'></td>                                                                                                                  \n");
                            htmlStr.Append("				  <td class=xl112 style='border-top:none'></td>                                                                                                                  \n");
                            htmlStr.Append("				  <td class=xl113 style='border-top:none'></td>                                                                                                                  \n");
                            htmlStr.Append("				  <td class=xl114>&nbsp;</td>                                                                                                                                    \n");
                            htmlStr.Append("		    </tr>                                                                                                                                                            \n");


                        }
                        else
                        {
                            htmlStr.Append("			 <tr class=xl115 height=53 style='mso-height-source:userset;height:" + v_rowHeightEmptyLast + "‬'>                                                                                         \n");
                            htmlStr.Append("				  <td height=53 class=xl106 style='height:" + v_rowHeightEmptyLast + "‬;border-right:none;border-top:none'>&nbsp;</td>                                                                 \n");
                            htmlStr.Append("				  <td class=xl107 width=32 style='width:30pt;border-top:none' x:num></td>                                                                                        \n");
                            htmlStr.Append("				  <td colspan=3  class=xl191 width=300 style='border-top:none;border-right:1.0pt solid black;                                                                     \n");
                            htmlStr.Append("				  width:225pt'></td>                                                                                                                                             \n");
                            htmlStr.Append("				  <td class=xl108 style='border-top:none'></td>                                                                                                                  \n");
                            htmlStr.Append("				  <td class=xl109 style='border-top:none' x:num></td>                                                                                                            \n");
                            htmlStr.Append("				  <td class=xl110 style='border-top:none' ></td>                                                                                                                 \n");
                            htmlStr.Append("				  <td class=xl111 style='border-top:none'></td>                                                                                                                  \n");
                            htmlStr.Append("				  <td class=xl112 style='border-top:none'></td>                                                                                                                  \n");
                            htmlStr.Append("				  <td class=xl113 style='border-top:none'></td>                                                                                                                  \n");
                            htmlStr.Append("				  <td class=xl114>&nbsp;</td>                                                                                                                                    \n");
                            htmlStr.Append("		    </tr>                                                                                                                                                            \n");

                        }
                    } // for

                }//Trang cuoi 11 dong

                if (k < v_countNumberOfPages - 1)
                {

                    htmlStr.Append(" 			 <tr class=xl115 height=53 style='mso-height-source:userset;height:" + (v_spacePerPage).ToString() + "pt'>                                                                                         \n");
                    htmlStr.Append("			  <td height=53 class=xl106 style='border-bottom:1.0pt solid black;height:" + (v_spacePerPage).ToString() + "pt;border-right:none;border-top:none'>&nbsp;</td>                                                                     \n");
                    htmlStr.Append("			  <td class=xl107 width=32 style='border-bottom:1.0pt solid black;width:30pt;border-top:none' x:num></td>                                                                                               \n");
                    htmlStr.Append("			  <td colspan=3  class=xl191 width=300 style='border-bottom:1.0pt solid black;border-right:1.0pt solid black;                                                                         \n");
                    htmlStr.Append("			  width:225pt;border-top:none'>&nbsp;</td>                                                                                                                                    \n");
                    htmlStr.Append("			  <td class=xl108 style='border-bottom:1.0pt solid black;border-top:none'>&nbsp;</td>                                                                                  \n");
                    htmlStr.Append("			  <td class=xl109 style='border-bottom:1.0pt solid black;border-top:none' > </td>    \n");
                    htmlStr.Append("			  <td class=xl110 style='border-bottom:1.0pt solid black;border-top:none' >&nbsp;</td>   \n");
                    htmlStr.Append("			  <td class=xl111 style='border-bottom:1.0pt solid black;border-top:none'></td>                                                                                        \n");
                    htmlStr.Append("			  <td class=xl112 style='border-bottom:1.0pt solid black;border-top:none'></td>                                                                                                                                              \n");
                    htmlStr.Append("			  <td class=xl113 style='border-bottom:1.0pt solid black;'></td>                                                                                                                                 \n");
                    htmlStr.Append("			  <td class=xl114 style='border-bottom:1.0pt solid black;'>&nbsp;</td>                                                                                                                                        \n");
                    htmlStr.Append(" 			</tr>                                                                                                                                                                \n");



                    htmlStr.Append("	<table  border=0>                                                                                                                                                                                                 \n");
                    htmlStr.Append("		<tr height=5  style='height: 55pt'>                                                                                                                                                                \n");
                    htmlStr.Append("		</tr>      																																														\n");
                    htmlStr.Append("	</table>             																																										\n");



                }


            }// for k                                                                                                                             

            //htmlStr.Append(" <tr class=xl144 height=21 style='mso-height-source:userset;height:19.6875pt'>                                                                                                     \n");
            //htmlStr.Append("  <td height=21 class=xl138 style='height:19.6875pt'>&nbsp;</td>                                                                                                                   \n");
            //htmlStr.Append("  <td class=xl139>*</td>                                                                                                                                                         \n");
            //htmlStr.Append("  <td colspan=3 class=xl188 width=300 style='border-right:1.0pt solid black;                                                                                                      \n");
            //htmlStr.Append("  width:225pt'>" + dt.Rows[0]["exchangerate_no"] + "</td>                                                                                                                                             \n");
            //htmlStr.Append("  <td class=xl140 width=54 style='border-left:none;width:51.25pt'>&nbsp;</td>                                                                                                       \n");
            //htmlStr.Append("  <td class=xl140 width=61 style='border-left:none;width:57.5pt'>&nbsp;</td>                                                                                                       \n");
            //htmlStr.Append("  <td class=xl141 width=33 style='width:31.25pt'>&nbsp;</td>                                                                                                                        \n");
            //htmlStr.Append("  <td class=xl141 width=82 style='width:77.5pt'>&nbsp;</td>                                                                                                                        \n");
            //htmlStr.Append("  <td class=xl141 width=12 style='width:11.25pt'>&nbsp;</td>                                                                                                                         \n");
            //htmlStr.Append("  <td class=xl142 width=124 style='width:116.25pt'>&nbsp;</td>                                                                                                                       \n");
            //htmlStr.Append("  <td class=xl143>&nbsp;</td>                                                                                                                                                    \n");
            //htmlStr.Append(" </tr>                                                                                                                                                                           \n");
            //htmlStr.Append("                                                                                                                                                                                 \n");
            //htmlStr.Append(" <tr height=20 style='mso-height-source:userset;height:18.75pt'>                                                                                                                  \n");
            //htmlStr.Append("  <td height=20 class=xl149 style='height:18.75pt'>&nbsp;</td>                                                                                                                    \n");
            //htmlStr.Append("  <td class=xl150></td>                                                                                                                                                          \n");
            //htmlStr.Append("  <td colspan=3 class=xl151 style='mso-ignore:colspan'></td>                                                                                                                     \n");
            //htmlStr.Append("  <td class=xl75 colspan=2 style='mso-ignore:colspan'>" + lb_amount_trans + "</td>												   \n");
            //htmlStr.Append("  <td class=xl153 colspan=2 style='text-align:left;vertical-align:middle;'>" + amount_trans + "</td>                          \n");
            //htmlStr.Append("  <td class=xl75 colspan=4 style='mso-ignore:colspan'>Cộng tiền hàng <i>(Net total)</i> :</font></td>                                                                            \n");
            //htmlStr.Append("  <td class=xl152></td>                                                                                                                                                          \n");
            //htmlStr.Append("  <td class=xl153>" + amount_net + " </td>                                                                                                                                           \n");
            //htmlStr.Append("  <td class=xl76>&nbsp;</td>                                                                                                                                                     \n");
            //htmlStr.Append(" </tr>                                                                                                                                                                           \n");

            htmlStr.Append("   <tr height=20 style='mso-height-source:userset;height:15.0pt'>																					\n");
            htmlStr.Append("    <td height=20 class=xl149 style='height:15.0pt;border-top: 0.5pt solid windowtext'>&nbsp;</td>                                                  \n");
            htmlStr.Append("    <td class=xl75 colspan=2 style='mso-ignore:colspan;border-top: 0.5pt solid windowtext'>" + lb_amount_trans + "</td>                                 \n");
            htmlStr.Append("    <td class=xl153 colspan=2 style='text-align:left;vertical-align:middle;border-top: 0.5pt solid windowtext;'>" + amount_trans + "</td>          \n");
            htmlStr.Append("    <td class=xl75 colspan=4 style='mso-ignore:colspan;border-top: 0.5pt solid windowtext'>Cộng tiền hàng <i>(Net total)</i> :</font></td>          \n");
            htmlStr.Append("    <td class=xl152 style='border-top: 0.5pt solid windowtext'></td>                                                                                \n");
            htmlStr.Append("    <td class=xl153 style='border-top: 0.5pt solid windowtext'>" + amount_net + "</td>                                                                 \n");
            htmlStr.Append("    <td class=xl76 style='border-top: 0.5pt solid windowtext'>&nbsp;</td>                                                                           \n");
            htmlStr.Append("   </tr>                                                                                                                                            \n");

            htmlStr.Append(" <tr height=20 style='mso-height-source:userset;height:18.75pt'>                                                                                                                  \n");
            htmlStr.Append("  <td height=20 class=xl149 style='height:18.75pt'>&nbsp;</td>                                                                                                                    \n");
            htmlStr.Append("  <td class=xl75 colspan=2 style='mso-ignore:colspan'>Thu&#7871; su&#7845;t                                                                                                      \n");
            htmlStr.Append("  GTGT (<font class='font45'>VAT rate</font><font class='font14'>)</font><font                                                                                                   \n");
            htmlStr.Append("  class='font46'> :</font>" + dt.Rows[0]["taxrate"] + "</td>                                                                                                                                  \n");  //" + dt.Rows[0]["TaxRate"] + "
            htmlStr.Append("  <td class=xl154></td>                                                                                                                                                          \n");
            htmlStr.Append("  <td class=xl150></td>                                                                                                                                                          \n");
            htmlStr.Append("  <td class=xl75 colspan=4 style='mso-ignore:colspan'>Ti&#7873;n thu&#7871;                                                                                                      \n");
            htmlStr.Append("  GTGT (<font class='font45'>VAT Amount</font><font class='font46'>) :</font></td>                                                                                               \n");
            htmlStr.Append("  <td class=xl152></td>                                                                                                                                                          \n");
            htmlStr.Append("  <td class=xl155>" + amount_vat + " </td>                                                                                                                                           \n");
            htmlStr.Append("  <td class=xl76>&nbsp;</td>                                                                                                                                                     \n");
            htmlStr.Append(" </tr>                                                                                                                                                                           \n");
            htmlStr.Append(" <tr height=20 style='mso-height-source:userset;height:18.75pt'>                                                                                                                  \n");
            htmlStr.Append("  <td height=20 class=xl149 style='height:18.75pt'>&nbsp;</td>                                                                                                                    \n");
            htmlStr.Append("  <td class=xl150></td>                                                                                                                                                          \n");
            htmlStr.Append("  <td colspan=3 class=xl151 style='mso-ignore:colspan'></td>                                                                                                                     \n");
            htmlStr.Append("  <td class=xl75 colspan=4 style='mso-ignore:colspan'>T&#7893;ng ti&#7873;n                                                                                                      \n");
            htmlStr.Append("  thanh toán (<font class='font45'>Gross total</font><font class='font46'>) :</font></td>                                                                                        \n");
            htmlStr.Append("  <td class=xl152></td>                                                                                                                                                          \n");
            htmlStr.Append("  <td class=xl155>" + amount_total + " </td>                                                                                                                                    \n");
            htmlStr.Append("  <td class=xl76>&nbsp;</td>                                                                                                                                                     \n");
            htmlStr.Append(" </tr>                                                                                                                                                                           \n");
            htmlStr.Append(" <tr class=xl144 height=21 style='mso-height-source:userset;height:15.95pt'>                                                                                                     \n");
            htmlStr.Append("  <td height=21 class=xl156 style='height:15.95pt'>&nbsp;</td>                                                                                                                   \n");
            htmlStr.Append("  <td class=xl75 colspan=3 style='mso-ignore:colspan'>S&#7889; ti&#7873;n                                                                                                        \n");
            htmlStr.Append("  b&#7857;ng ch&#7919; (<i>Amount in words</i><font                                                                                                                              \n");
            htmlStr.Append("  class='font14'>) :</font></td>                                                                                                                                                 \n");
            htmlStr.Append("  <td colspan=7 rowspan=2 class=xl185 width=427 style='width:322pt'>" + read_prive + " </td>                                                                                        \n");
            htmlStr.Append("  <td class=xl158>&nbsp;</td>                                                                                                                                                    \n");
            htmlStr.Append(" </tr>                                                                                                                                                                           \n");
            htmlStr.Append(" <tr class=xl144 height=21 style='mso-height-source:userset;height:15.95pt'>                                                                                                     \n");
            htmlStr.Append("  <td height=21 class=xl156 style='height:15.95pt'>&nbsp;</td>                                                                                                                   \n");
            htmlStr.Append("  <td class=xl159></td>                                                                                                                                                          \n");
            htmlStr.Append("  <td colspan=2 class=xl160 style='mso-ignore:colspan'></td>                                                                                                                     \n");
            htmlStr.Append("  <td class=xl158>&nbsp;</td>                                                                                                                                                    \n");
            htmlStr.Append(" </tr>                                                                                                                                                                           \n");
            htmlStr.Append(" <tr class=xl167 height=31																					  \n");
            htmlStr.Append(" 	style='mso-height-source: userset; height: 23.25pt'>                                                      \n");
            htmlStr.Append(" 	<td height=31 class=xl165                                                                                 \n");
            htmlStr.Append(" 		style='height: 23.25pt; border-top: .5pt solid black;'>&nbsp;</td>                                    \n");
            htmlStr.Append(" 	<td colspan=2 class=xl186 width=191                                                                       \n");
            htmlStr.Append(" 		style='width: 143pt; border-top: .5pt solid black;'>Ng&#432;&#7901;i                                  \n");
            htmlStr.Append(" 		th&#7921;c hi&#7879;n chuy&#7875;n &#273;&#7893;i (Converter)</td>                                    \n");
            htmlStr.Append(" 	<td colspan=2 class=xl186 width=141                                                                       \n");
            htmlStr.Append(" 		style='width: 106pt; border-top: .5pt solid black;'>Ng&#432;&#7901;i                                  \n");
            htmlStr.Append(" 		mua (Buyer)</td>                                                                                      \n");
            htmlStr.Append(" 	<td colspan=3 class=xl186 width=148                                                                       \n");
            htmlStr.Append(" 		style='width: 112pt; border-top: .5pt solid black;'></td>                                             \n");
            htmlStr.Append(" 	<td colspan=3 class=xl187 width=218                                                                       \n");
            htmlStr.Append(" 		style='width: 164pt; border-top: .5pt solid black;'>Người bán                                         \n");
            htmlStr.Append(" 		(Seller)</td>                                                                                         \n");
            htmlStr.Append(" 	<td class=xl166 style='border-top: .5pt solid black;'>&nbsp;</td>                                         \n");
            htmlStr.Append(" </tr>                                                                                                       \n");
            htmlStr.Append(" <tr height=17 style='height: 12.75pt'>                                                                      \n");
            htmlStr.Append(" 	<td height=17 class=xl71 style='height: 12.75pt'>&nbsp;</td>                                              \n");
            htmlStr.Append(" 	<td colspan=2 class=xl177 width=191 style='width: 143pt'>Ký,                                              \n");
            htmlStr.Append(" 		ghi rõ h&#7885; tên</td>                                                                              \n");
            htmlStr.Append(" 	<td colspan=2 class=xl177 width=141 style='width: 106pt'>Ký,                                              \n");
            htmlStr.Append(" 		ghi rõ h&#7885; tên</td>                                                                              \n");
            htmlStr.Append(" 	<td colspan=3 class=xl177 width=148 style='width: 112pt'></td>                                            \n");
            htmlStr.Append(" 	<td colspan=3 class=xl177 width=218 style='width: 164pt'>Ký,                                              \n");
            htmlStr.Append(" 		&#273;óng d&#7845;u, ghi rõ h&#7885; tên</td>                                                         \n");
            htmlStr.Append(" 	<td class=xl76>&nbsp;</td>                                                                                \n");
            htmlStr.Append(" </tr>                                                                                                       \n");
            htmlStr.Append(" <tr height=17 style='height: 12.75pt'>                                                                      \n");
            htmlStr.Append(" 	<td height=17 class=xl71 style='height: 12.75pt'>&nbsp;</td>                                              \n");
            htmlStr.Append(" 	<td colspan=2 class=xl177 width=191 style='width: 143pt'>(Sign                                            \n");
            htmlStr.Append(" 		&amp; full name)</td>                                                                                 \n");
            htmlStr.Append(" 	<td colspan=2 class=xl177 width=141 style='width: 106pt'>(Sign                                            \n");
            htmlStr.Append(" 		&amp; full name)</td>                                                                                 \n");
            htmlStr.Append(" 	<td colspan=3 class=xl177 width=148 style='width: 112pt'></td>                                            \n");
            htmlStr.Append(" 	<td colspan=3 class=xl177 width=218 style='width: 164pt'>(Sign,                                           \n");
            htmlStr.Append(" 		stamp &amp; full name)</td>                                                                           \n");
            htmlStr.Append(" 	<td class=xl76>&nbsp;</td>                                                                                \n");
            htmlStr.Append(" </tr>                                                                                                       \n");
            htmlStr.Append(" <tr height=20 style='mso-height-source:userset;height:18.75pt'>                                                                                                                  \n");
            htmlStr.Append("  <td height=20 class=xl71 style='height:18.75pt'>&nbsp;</td>                                                                                                                     \n");
            htmlStr.Append("  <td class=xl169></td>                                                                                                                                                          \n");
            htmlStr.Append("  <td colspan=9 class=xl170 style='mso-ignore:colspan'></td>                                                                                                                     \n");
            htmlStr.Append("  <td class=xl76>&nbsp;</td>                                                                                                                                                     \n");
            htmlStr.Append(" </tr>                                                                                                                                                                           \n");
            htmlStr.Append(" <tr height=20 style='mso-height-source:userset;height:10.75pt'>                                                                                                                  \n");
            htmlStr.Append("  <td height=20 class=xl71 style='height:10.75pt'>&nbsp;</td>                                                                                                                     \n");
            htmlStr.Append("  <td class=xl169></td>                                                                                                                                                          \n");
            htmlStr.Append("  <td colspan=9 class=xl170 style='mso-ignore:colspan'></td>                                                                                                                     \n");
            htmlStr.Append("  <td class=xl76>&nbsp;</td>                                                                                                                                                     \n");
            htmlStr.Append(" </tr>                                                                                                                                                                           \n");

            if (dt.Rows[0]["sign_yn"].ToString() == "Y")
            {

                htmlStr.Append("  <tr height=20 style='mso-height-source:userset;height:18.75pt'>                                                                                                                 \n");
                htmlStr.Append("  <td height=20 class=xl71 style='height:18.75pt;border-right:none;'></td>                                                                                                        \n");
                htmlStr.Append("  <td class=xl167></td>                                                                                                                                                          \n");
                htmlStr.Append("  <td colspan=4 class=xl168 style='mso-ignore:colspan;border-left:none;'></td>                                                                                                   \n");
                htmlStr.Append("  <td colspan=5 height=20 width=312 style='height:18.75pt;width:235pt' align=left valign=top>                                                                                     \n");
                htmlStr.Append("  <span style='mso-ignore:vglayout;position:                                                                                                                                     \n");
                htmlStr.Append("  absolute;z-index:4;margin-left:117px;margin-top:9px;width:62px;height:42px'><img                                                                                               \n");
                htmlStr.Append("  width=62 height=42  src='D:/webproject/e-invoice-ws/02.Web/EInvoice/img/check_signed.png'  v:shapes='Rectangle_x0020_4'></span><![endif]></td>                                  \n");
                htmlStr.Append("  <td class=xl76>&nbsp;</td>                                                                                                                                                     \n");
                htmlStr.Append(" </tr>                                                                                                                                                                           \n");

            }

            htmlStr.Append("   <tr height=20 style='mso-height-source:userset;height:18.75pt'>																																											\n");
            htmlStr.Append("    <td height=20 class=xl71 style='height:18.75pt'>&nbsp;" + dt.Rows[0]["convert_name"] + "</td>                                                                                                                                                                             \n");
            htmlStr.Append("    <td class=xl169></td>                                                                                                                                                                                                                  \n");
            htmlStr.Append("    <td colspan=4 class=xl170 style='mso-ignore:colspan'></td>                                                                                                                                                                             \n");
            htmlStr.Append("    <td colspan=5 class=xl178 width=312 style='border-right:.5pt solid #0066CC;                                                                                                                                                            \n");
            htmlStr.Append("    width:235pt' x:str='Signature Valid'>Signature Valid<span                                                                                                                                                                             \n");
            htmlStr.Append("    style='mso-spacerun:yes'> </span></td>                                                                                                                                                                                                 \n");
            htmlStr.Append("    <td class=xl76>&nbsp;</td>                                                                                                                                                                                                             \n");
            htmlStr.Append("   </tr>                                                                                                                                                                                                                                   \n");
            htmlStr.Append("   <tr height=36 style='mso-height-source:userset;height:27.0pt'>                                                                                                                                                                          \n");
            htmlStr.Append("    <td height=36 class=xl71 style='height:27.0pt'>&nbsp;" + dt.Rows[0]["convert_date"] + "</td>                                                                                                                                                                             \n");
            htmlStr.Append("    <td colspan=2 class=xl181></td>                                                                                                                                                                                                        \n");
            htmlStr.Append("    <td colspan=3 class=xl170 style='mso-ignore:colspan'></td>                                                                                                                                                                             \n");
            htmlStr.Append("    <td colspan=5 class=xl182 width=312 style='border-right:.5pt solid #0066CC;                                                                                                                                                            \n");
            htmlStr.Append("    width:235pt'>&#272;&#432;&#7907;c ký b&#7903;i:<font class='font64'>" + dt.Rows[0]["SignedBy"] + "</font></td>                                                                                                                                      \n");
            htmlStr.Append("    <td class=xl76>&nbsp;</td>                                                                                                                                                                                                             \n");
            htmlStr.Append("   </tr>                                                                                                                                                                                                                                   \n");
            htmlStr.Append("   <tr height=20 style='mso-height-source:userset;height:18.75pt'>                                                                                                                                                                          \n");
            htmlStr.Append("    <td height=20 class=xl71 style='height:18.75pt'>&nbsp;</td>                                                                                                                                                                             \n");
            htmlStr.Append("    <td class=xl170 colspan=5 style='mso-ignore:colspan'>Mã CQT: <font style='color: #003366;font-weight: 400;font-family:'Times New Roman', serif;font-size: 12pt'>&nbsp;&nbsp;" + dt.Rows[0]["cqt_mccqt_id"] + "</font></td>                                   \n");
            htmlStr.Append("    <td colspan=5 class=xl173 style='border-right:.5pt solid #0066CC'><font                                                                                                                                                                \n");
            htmlStr.Append("    class='font18'>Ngày ký</font><font class='font64'>: " + dt.Rows[0]["SignedDate"] + "</font></td>                                                                                                                                                         \n");
            htmlStr.Append("    <td class=xl76>&nbsp;</td>                                                                                                                                                                                                             \n");
            htmlStr.Append("   </tr>                                                                                                                                                                                                                                   \n");
            htmlStr.Append("   <tr height=20 style='mso-height-source:userset;height:18.75pt'>                                                                                                                                                                           \n");
            htmlStr.Append("    <td height=20 class=xl71 style='height:18.75pt'>&nbsp;</td>                                                                                                                                                                             \n");
            htmlStr.Append("   <td class=xl171 colspan=4 style='mso-ignore:colspan'><a                                                                                                                                                                                 \n");
            htmlStr.Append("    href=" + dt.Rows[0]["WEBSITE_EI"] + "><span style='color:#003366;                                                                                                                                                                                       \n");
            htmlStr.Append("    font-style:italic'>" + dt.Rows[0]["WEBSITE_EI"] + "</span></a></td>                                                                                                                                                                                     \n");
            htmlStr.Append("     <td class=xl76 colspan=1 style='border-right: none'>&nbsp;</td>                                                                                                                                                                       \n");
            htmlStr.Append("    <td colspan=5 class=xl170 style='mso-ignore:colspan'>Mã nh&#7853;n hóa &#273;&#417;n: <font style='color: #003366;font-weight: 400;font-family:'Times New Roman', serif;font-size: 12pt'>&nbsp;&nbsp;" + dt.Rows[0]["matracuu"] + "</font></td>      \n");
            htmlStr.Append("    <td class=xl76>&nbsp;</td>                                                                                                                                                                                                             \n");
            htmlStr.Append("   </tr>                                                                                                                                                                                                                                   \n");
            htmlStr.Append("   <tr height=22 style='mso-height-source:userset;height:17.1pt'>                                                                                                                                                                          \n");
            htmlStr.Append("    <td height=22 class=xl83 style='height:17.1pt'>&nbsp;</td>                                                                                                                                                                             \n");
            htmlStr.Append("    <td colspan=10 class=xl176>(C&#7847;n ki&#7875;m tra &#273;&#7889;i                                                                                                                                                                    \n");
            htmlStr.Append("    chi&#7871;u khi l&#7853;p, giao, nh&#7853;n hóa &#273;&#417;n)</td>                                                                                                                                                                    \n");
            htmlStr.Append("    <td class=xl89>&nbsp;</td>                                                                                                                                                                                                             \n");
            htmlStr.Append("   </tr>                                                                                                                                                                                                                                   \n");


            htmlStr.Append(" <![if supportMisalignedColumns]>                                                                                                                                                \n");
            htmlStr.Append(" <tr height=0 style='display:none'>                                                                                                                                              \n");
            htmlStr.Append("  <td width=7 style='width:6.25pt'></td>                                                                                                                                            \n");
            htmlStr.Append("  <td width=32 style='width:30pt'></td>                                                                                                                                          \n");
            htmlStr.Append("  <td width=159 style='width:148.75pt'></td>                                                                                                                                        \n");
            htmlStr.Append("  <td width=80 style='width:75pt'></td>                                                                                                                                          \n");
            htmlStr.Append("  <td width=61 style='width:57.5pt'></td>                                                                                                                                          \n");
            htmlStr.Append("  <td width=54 style='width:51.25pt'></td>                                                                                                                                          \n");
            htmlStr.Append("  <td width=61 style='width:57.5pt'></td>                                                                                                                                          \n");
            htmlStr.Append("  <td width=33 style='width:31.25pt'></td>                                                                                                                                          \n");
            htmlStr.Append("  <td width=82 style='width:77.5pt'></td>                                                                                                                                          \n");
            htmlStr.Append("  <td width=12 style='width:11.25pt'></td>                                                                                                                                           \n");
            htmlStr.Append("  <td width=124 style='width:116.25pt'></td>                                                                                                                                         \n");
            htmlStr.Append("  <td width=8 style='width:7.5pt'></td>                                                                                                                                            \n");
            htmlStr.Append(" </tr>                                                                                                                                                                           \n");
            htmlStr.Append(" <![endif]>                                                                                                                                                                      \n");
            htmlStr.Append("</table>                                                                                                                                                                         \n");
            htmlStr.Append("</section>                                                                                                                                                                       \n");
            htmlStr.Append("</body>                                                                                                                                                                          \n");
            htmlStr.Append("</html>                                                                                                                                                                          \n");


            string filePath = "";

            using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"D:\webproject\e-invoice-ws\02.Web\AttachFileText\" + tei_einvoice_m_pk + ".html"))
            {
                file.WriteLine(htmlStr.ToString()); // "sb" is the StringBuilder
            }

            connection.Close();
            connection.Dispose();
            return htmlStr.ToString() + "|" + dt.Rows[0]["templateCode"].ToString().Replace("/", "") + "_" + dt.Rows[0]["InvoiceSerialNo"].ToString().Replace("/", "") + "_" + dt.Rows[0]["InvoiceNumber"];

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
            int max_length = 100;// 42;
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