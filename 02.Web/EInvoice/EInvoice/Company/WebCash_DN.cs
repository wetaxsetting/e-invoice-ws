
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
    public class WebCash_DN
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


            

            //read_prive = read_prive.Substring(0, 2) + read_prive.Substring(2, read_prive.Length - 2).ToLower();
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
            htmlStr.Append("	<!--table                                                                                                                                                                    \n");
            htmlStr.Append("		{mso-displayed-decimal-separator:'\\.';                                                                                                                                   \n");
            htmlStr.Append("		mso-displayed-thousand-separator:'\\, ';}                                                                                                                                  \n");
            htmlStr.Append("	.font522214                                                                                                                                                                  \n");
            htmlStr.Append("		{color:windowtext;                                                                                                                                                       \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;}                                                                                                                                                     \n");
            htmlStr.Append("	.font622214                                                                                                                                                                  \n");
            htmlStr.Append("		{color:windowtext;                                                                                                                                                       \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:italic;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;}                                                                                                                                                     \n");
            htmlStr.Append("	.font722214                                                                                                                                                                  \n");
            htmlStr.Append("		{color:windowtext;                                                                                                                                                       \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;}                                                                                                                                                     \n");
            htmlStr.Append("	.font822214                                                                                                                                                                  \n");
            htmlStr.Append("		{color:windowtext;                                                                                                                                                       \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;}                                                                                                                                                     \n");
            htmlStr.Append("	.font922214                                                                                                                                                                  \n");
            htmlStr.Append("		{color:#993300;                                                                                                                                                          \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;}                                                                                                                                                     \n");
            htmlStr.Append("	.font1022214                                                                                                                                                                 \n");
            htmlStr.Append("		{color:windowtext;                                                                                                                                                       \n");
            htmlStr.Append("		font-size:13.2pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;}                                                                                                                                                     \n");
            htmlStr.Append("	.font1122214                                                                                                                                                                 \n");
            htmlStr.Append("		{color:black;                                                                                                                                                            \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;}                                                                                                                                                     \n");
            htmlStr.Append("	.font1222214                                                                                                                                                                 \n");
            htmlStr.Append("		{color:#333399;                                                                                                                                                          \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;}                                                                                                                                                     \n");
            htmlStr.Append("	.font1322214                                                                                                                                                                 \n");
            htmlStr.Append("		{color:windowtext;                                                                                                                                                       \n");
            htmlStr.Append("		font-size:13.2pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:italic;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;}                                                                                                                                                     \n");
            htmlStr.Append("	.font1422214                                                                                                                                                                 \n");
            htmlStr.Append("		{color:black;                                                                                                                                                            \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;}                                                                                                                                                     \n");
            htmlStr.Append("	.font1522214                                                                                                                                                                 \n");
            htmlStr.Append("		{color:#002060;                                                                                                                                                          \n");
            htmlStr.Append("		font-size:13.2pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;}                                                                                                                                                     \n");
            htmlStr.Append("	.font1622214                                                                                                                                                                 \n");
            htmlStr.Append("		{color:red;                                                                                                                                                              \n");
            htmlStr.Append("		font-size:11.25pt;                                                                                                                                                         \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;}                                                                                                                                                     \n");
            htmlStr.Append("	.font1722214                                                                                                                                                                 \n");
            htmlStr.Append("		{color:black;                                                                                                                                                            \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;}                                                                                                                                                     \n");
            htmlStr.Append("	.font1822214                                                                                                                                                                 \n");
            htmlStr.Append("		{color:black;                                                                                                                                                            \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:italic;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;}                                                                                                                                                     \n");
            htmlStr.Append("	.font1922214                                                                                                                                                                 \n");
            htmlStr.Append("		{color:windowtext;                                                                                                                                                       \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:italic;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;}                                                                                                                                                     \n");
            htmlStr.Append("	.font2022214                                                                                                                                                                 \n");
            htmlStr.Append("		{color:#002060;                                                                                                                                                          \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;}                                                                                                                                                     \n");
            htmlStr.Append("	.font2122214                                                                                                                                                                 \n");
            htmlStr.Append("		{color:black;                                                                                                                                                            \n");
            htmlStr.Append("		font-size:13.2pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl6322214                                                                                                                                                                   \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl6422214                                                                                                                                                                   \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:15.0pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl6522214                                                                                                                                                                   \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:16.25pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl6622214                                                                                                                                                                   \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:16.25pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl6722214                                                                                                                                                                   \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:#C00000;                                                                                                                                                           \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl6822214                                                                                                                                                                   \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl6922214                                                                                                                                                                   \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl7022214                                                                                                                                                                   \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:italic;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl7122214                                                                                                                                                                   \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:1.0pt;                                                                                                                                                         \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl7222214                                                                                                                                                                   \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:1.0pt;                                                                                                                                                         \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl7322214                                                                                                                                                                   \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl7422214                                                                                                                                                                   \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:1.0pt;                                                                                                                                                         \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl7522214                                                                                                                                                                   \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl7622214                                                                                                                                                                   \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl7722214                                                                                                                                                                   \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl7822214                                                                                                                                                                   \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl7922214                                                                                                                                                                   \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl8022214                                                                                                                                                                   \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl8122214                                                                                                                                                                   \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:16.25pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl8222214                                                                                                                                                                   \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:16.25pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl8322214                                                                                                                                                                   \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:left;                                                                                                                                                         \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl8422214                                                                                                                                                                   \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl8522214                                                                                                                                                                   \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl8622214                                                                                                                                                                   \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl8722214                                                                                                                                                                   \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:17.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl8822214                                                                                                                                                                   \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:red;                                                                                                                                                               \n");
            htmlStr.Append("		font-size:21.25pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:Arial, sans-serif;                                                                                                                                           \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:justify;                                                                                                                                                  \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl8922214                                                                                                                                                                   \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:Arial, sans-serif;                                                                                                                                           \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl9022214                                                                                                                                                                   \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:Arial, sans-serif;                                                                                                                                           \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl9122214                                                                                                                                                                   \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:Arial, sans-serif;                                                                                                                                           \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl9222214                                                                                                                                                                   \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:Arial, sans-serif;                                                                                                                                           \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl9322214                                                                                                                                                                   \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:italic;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:Arial, sans-serif;                                                                                                                                           \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl9422214                                                                                                                                                                   \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:Arial, sans-serif;                                                                                                                                           \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl9522214                                                                                                                                                                   \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:11.25pt;                                                                                                                                                         \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl9622214                                                                                                                                                                   \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:black;                                                                                                                                                             \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl9722214                                                                                                                                                                   \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:Arial, sans-serif;                                                                                                                                           \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:left;                                                                                                                                                         \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl9822214                                                                                                                                                                   \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl9922214                                                                                                                                                                   \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl10022214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl10122214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl10222214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:red;                                                                                                                                                               \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl10322214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:red;                                                                                                                                                               \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl10422214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:16.25pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl10522214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:red;                                                                                                                                                               \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:italic;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl10622214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:15.0pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:Arial, sans-serif;                                                                                                                                           \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl10722214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:15.0pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:Arial, sans-serif;                                                                                                                                           \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl10822214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:15.0pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:Arial, sans-serif;                                                                                                                                           \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl10922214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl11022214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl11122214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:1.0pt dotted windowtext;                                                                                                                                       \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:1.0pt dotted windowtext;                                                                                                                                    \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl11222214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:1.0pt dotted windowtext;                                                                                                                                       \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:1.0pt dotted windowtext;                                                                                                                                    \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl11322214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl11422214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:1.0pt dotted windowtext;                                                                                                                                       \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:1.0pt dotted windowtext;                                                                                                                                    \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl11522214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:1.0pt dotted windowtext;                                                                                                                                       \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:1.0pt dotted windowtext;                                                                                                                                    \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl11622214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:1.0pt dotted windowtext;                                                                                                                                       \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:1.0pt dotted windowtext;                                                                                                                                    \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl11722214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:red;                                                                                                                                                               \n");
            htmlStr.Append("		font-size:22.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;                                                                                                                                                      \n");
            htmlStr.Append("		mso-text-control:shrinktofit;}                                                                                                                                           \n");
            htmlStr.Append("	.xl11822214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:red;                                                                                                                                                               \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl11922214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:red;                                                                                                                                                               \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl12022214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:black;                                                                                                                                                             \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl12122214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:red;                                                                                                                                                               \n");
            htmlStr.Append("		font-size:21.25pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:justify;                                                                                                                                                  \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl12222214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:red;                                                                                                                                                               \n");
            htmlStr.Append("		font-size:18.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:justify;                                                                                                                                                  \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl12322214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl12422214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl12522214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.2pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:left;                                                                                                                                                         \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl12622214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.2pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:left;                                                                                                                                                         \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl12722214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:right;                                                                                                                                                        \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl12822214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:black;                                                                                                                                                             \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:right;                                                                                                                                                        \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl12922214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:Arial, sans-serif;                                                                                                                                           \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:top;                                                                                                                                                      \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl13022214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:top;                                                                                                                                                      \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl130222141                                                                                                                                                                 \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:top;                                                                                                                                                      \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl13122214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:Arial, sans-serif;                                                                                                                                           \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:top;                                                                                                                                                      \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl13222214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:Arial, sans-serif;                                                                                                                                           \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:top;                                                                                                                                                      \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl13322214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:top;                                                                                                                                                      \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;                                                                                                                                                      \n");
            htmlStr.Append("		mso-text-control:shrinktofit;}                                                                                                                                           \n");
            htmlStr.Append("	.xl13422214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:left;                                                                                                                                                         \n");
            htmlStr.Append("		vertical-align:top;                                                                                                                                                      \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl13522214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:top;                                                                                                                                                      \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl13622214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:top;                                                                                                                                                      \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl13722214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl13822214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl13922214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl14022214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl14122214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl14222214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl14322214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:left;                                                                                                                                                         \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl14422214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:italic;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl14522214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:red;                                                                                                                                                               \n");
            htmlStr.Append("		font-size:20.0pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Arial Rounded MT Bold', sans-serif;                                                                                                                         \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:right;                                                                                                                                                        \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl14622214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:red;                                                                                                                                                               \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:italic;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:justify;                                                                                                                                                  \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl14722214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:Arial, sans-serif;                                                                                                                                           \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl14822214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:Arial, sans-serif;                                                                                                                                           \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl14922214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:Arial, sans-serif;                                                                                                                                           \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl15022214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:Arial, sans-serif;                                                                                                                                           \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl15122214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:black;                                                                                                                                                             \n");
            htmlStr.Append("		font-size:13.2pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:right;                                                                                                                                                        \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl15222214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                           \n");
            htmlStr.Append("		text-align:right;                                                                                                                                                        \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:1.0pt dotted windowtext;                                                                                                                                       \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:1.0pt dotted windowtext;                                                                                                                                    \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl15322214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                           \n");
            htmlStr.Append("		text-align:right;                                                                                                                                                        \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:1.0pt dotted windowtext;                                                                                                                                       \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:1.0pt dotted windowtext;                                                                                                                                    \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl15422214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                           \n");
htmlStr.Append("		text-align:right;                                                                                                                                                        \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:1.0pt dotted windowtext;                                                                                                                                       \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:1.0pt dotted windowtext;                                                                                                                                    \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl15522214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                           \n");
htmlStr.Append("		text-align:right;                                                                                                                                                        \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:1.0pt dotted windowtext;                                                                                                                                       \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl15622214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                           \n");
htmlStr.Append("		text-align:right;                                                                                                                                                        \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:1.0pt dotted windowtext;                                                                                                                                       \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl15722214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                           \n");
htmlStr.Append("		text-align:right;                                                                                                                                                        \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:1.0pt dotted windowtext;                                                                                                                                       \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl15822214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:right;                                                                                                                                                        \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl15922214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:right;                                                                                                                                                        \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl16022214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:right;                                                                                                                                                        \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl16122214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:left;                                                                                                                                                         \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl16222214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl16322214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl16422214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:red;                                                                                                                                                               \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:italic;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:right;                                                                                                                                                        \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl16522214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl16622214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl16722214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl16822214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl16922214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl17022214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl17122214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl17222214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl17322214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl17422214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                           \n");
htmlStr.Append("		text-align:right;                                                                                                                                                        \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:1.0pt dotted windowtext;                                                                                                                                    \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl17522214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                           \n");
htmlStr.Append("		text-align:right;                                                                                                                                                        \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:1.0pt dotted windowtext;                                                                                                                                    \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl17622214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                           \n");
htmlStr.Append("		text-align:right;                                                                                                                                                        \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:1.0pt dotted windowtext;                                                                                                                                    \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl17722214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.2pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:left;                                                                                                                                                         \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;                                                                                                                                                      \n");
            htmlStr.Append("		mso-text-control:shrinktofit;}                                                                                                                                           \n");
            htmlStr.Append("	.xl17822214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.2pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:left;                                                                                                                                                         \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;                                                                                                                                                      \n");
            htmlStr.Append("		mso-text-control:shrinktofit;}                                                                                                                                           \n");
            htmlStr.Append("	.xl17922214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.2pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:left;                                                                                                                                                         \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl18022214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.2pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:left;                                                                                                                                                         \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl18122214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl18222214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:#002060;                                                                                                                                                           \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:left;                                                                                                                                                         \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl18322214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                           \n");
htmlStr.Append("		text-align:right;                                                                                                                                                        \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:1.0pt dotted windowtext;                                                                                                                                       \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:1.0pt dotted windowtext;                                                                                                                                    \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl18422214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:1.0pt dotted windowtext;                                                                                                                                       \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:1.0pt dotted windowtext;                                                                                                                                    \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl18522214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:1.0pt dotted windowtext;                                                                                                                                       \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:1.0pt dotted windowtext;                                                                                                                                    \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl18622214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:1.0pt dotted windowtext;                                                                                                                                       \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:1.0pt dotted windowtext;                                                                                                                                    \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl18722214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                           \n");
htmlStr.Append("		text-align:right;                                                                                                                                                        \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:1.0pt dotted windowtext;                                                                                                                                       \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:1.0pt dotted windowtext;                                                                                                                                    \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl18822214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                           \n");
htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:1.0pt dotted windowtext;                                                                                                                                       \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:1.0pt dotted windowtext;                                                                                                                                    \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl18922214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl19022214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl19122214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border:.5pt solid windowtext;                                                                                                                                            \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl19222214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:red;                                                                                                                                                               \n");
            htmlStr.Append("		font-size:22.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;                                                                                                                                                      \n");
            htmlStr.Append("		mso-text-control:shrinktofit;}                                                                                                                                           \n");
            htmlStr.Append("	.xl19322214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:red;                                                                                                                                                               \n");
            htmlStr.Append("		font-size:22.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;                                                                                                                                                      \n");
            htmlStr.Append("		mso-text-control:shrinktofit;}                                                                                                                                           \n");
            htmlStr.Append("	.xl19422214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:red;                                                                                                                                                               \n");
            htmlStr.Append("		font-size:22.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;                                                                                                                                                      \n");
            htmlStr.Append("		mso-text-control:shrinktofit;}                                                                                                                                           \n");
            htmlStr.Append("	.xl19522214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:red;                                                                                                                                                               \n");
            htmlStr.Append("		font-size:22.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;                                                                                                                                                      \n");
            htmlStr.Append("		mso-text-control:shrinktofit;}                                                                                                                                           \n");
            htmlStr.Append("	.xl19622214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl19722214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl19822214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                           \n");
            htmlStr.Append("		text-align:right;                                                                                                                                                        \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl19922214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                           \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl20022214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                           \n");
htmlStr.Append("		text-align:right;                                                                                                                                                        \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl20122214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                           \n");
htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:1.0pt dotted windowtext;                                                                                                                                       \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:1.0pt dotted windowtext;                                                                                                                                    \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl20222214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:1.0pt dotted windowtext;                                                                                                                                       \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:1.0pt dotted windowtext;                                                                                                                                    \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl20322214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:1.0pt dotted windowtext;                                                                                                                                       \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:1.0pt dotted windowtext;                                                                                                                                    \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl20422214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                           \n");
htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:1.0pt dotted windowtext;                                                                                                                                       \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:1.0pt dotted windowtext;                                                                                                                                    \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;                                                                                                                                                      \n");
            htmlStr.Append("		mso-text-control:shrinktofit;}                                                                                                                                           \n");
            htmlStr.Append("	.xl20522214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                           \n");
htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:1.0pt dotted windowtext;                                                                                                                                       \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:1.0pt dotted windowtext;                                                                                                                                    \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;                                                                                                                                                      \n");
            htmlStr.Append("		mso-text-control:shrinktofit;}                                                                                                                                           \n");
            htmlStr.Append("	.xl20622214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                           \n");
htmlStr.Append("		text-align:right;                                                                                                                                                        \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:1.0pt dotted windowtext;                                                                                                                                       \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:1.0pt dotted windowtext;                                                                                                                                    \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;                                                                                                                                                      \n");
            htmlStr.Append("		mso-text-control:shrinktofit;}                                                                                                                                           \n");
            htmlStr.Append("	.xl20722214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:red;                                                                                                                                                               \n");
            htmlStr.Append("		font-size:18.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:justify;                                                                                                                                                  \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl20822214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:italic;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl20922214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:black;                                                                                                                                                             \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:left;                                                                                                                                                         \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl21022214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:black;                                                                                                                                                             \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:left;                                                                                                                                                         \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl21122214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:black;                                                                                                                                                             \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:left;                                                                                                                                                         \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl21222214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:red;                                                                                                                                                               \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:left;                                                                                                                                                         \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl21322214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:red;                                                                                                                                                               \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:left;                                                                                                                                                         \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl21422214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:red;                                                                                                                                                               \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:left;                                                                                                                                                         \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl21522214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:11.25pt;                                                                                                                                                         \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:italic;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl21622214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:11.25pt;                                                                                                                                                         \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:italic;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl21722214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:11.25pt;                                                                                                                                                         \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:italic;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl21822214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:11.25pt;                                                                                                                                                         \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:italic;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl21922214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                           \n");
htmlStr.Append("		text-align:right;                                                                                                                                                        \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:1.0pt dotted windowtext;                                                                                                                                    \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl22022214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl22122214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border:.5pt solid windowtext;                                                                                                                                            \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl22222214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl22322214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:13.75pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:700;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:'\\#\\,\\#\\#0\\;\\[Red\\]\\#\\,\\#\\#0';                                                                                                                         \n");
            htmlStr.Append("		text-align:right;                                                                                                                                                        \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border:.5pt solid windowtext;                                                                                                                                            \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl22422214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl22522214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl22622214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl22722214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl22822214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl22922214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                           \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl23022214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                           \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl23122214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                           \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:1.0pt dotted windowtext;                                                                                                                                       \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:1.0pt dotted windowtext;                                                                                                                                    \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;                                                                                                                                                      \n");
            htmlStr.Append("		mso-text-control:shrinktofit;}                                                                                                                                           \n");
            htmlStr.Append("	.xl23222214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:left;                                                                                                                                                         \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:1.0pt dotted windowtext;                                                                                                                                       \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:1.0pt dotted windowtext;                                                                                                                                    \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl23322214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:left;                                                                                                                                                         \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:1.0pt dotted windowtext;                                                                                                                                       \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:1.0pt dotted windowtext;                                                                                                                                    \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl23422214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:12.5pt;                                                                                                                                                        \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:left;                                                                                                                                                         \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:1.0pt dotted windowtext;                                                                                                                                       \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:1.0pt dotted windowtext;                                                                                                                                    \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl23522214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:1.0pt;                                                                                                                                                         \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl23622214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:1.0pt;                                                                                                                                                         \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:left;                                                                                                                                                         \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl23722214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:1.0pt;                                                                                                                                                         \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:left;                                                                                                                                                         \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl23822214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:1.0pt;                                                                                                                                                         \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl23922214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:1.0pt;                                                                                                                                                         \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl24022214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:1.0pt;                                                                                                                                                         \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl24122214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:1.0pt;                                                                                                                                                         \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl24222214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:1.0pt;                                                                                                                                                         \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:Arial, sans-serif;                                                                                                                                           \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:left;                                                                                                                                                         \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl24322214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:1.0pt;                                                                                                                                                         \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:Arial, sans-serif;                                                                                                                                           \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:left;                                                                                                                                                         \n");
            htmlStr.Append("		vertical-align:top;                                                                                                                                                      \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:normal;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl24422214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:1.0pt;                                                                                                                                                         \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:middle;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl24522214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:1.0pt;                                                                                                                                                         \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:.5pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl24622214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:1.0pt;                                                                                                                                                         \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:none;                                                                                                                                                         \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl24722214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:1.0pt;                                                                                                                                                         \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:.5pt solid windowtext;                                                                                                                                       \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl24822214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:1.0pt;                                                                                                                                                         \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl24922214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:1.0pt;                                                                                                                                                         \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:'\\@';                                                                                                                                                  \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl25022214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:1.0pt;                                                                                                                                                         \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl25122214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:1.0pt;                                                                                                                                                         \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:'\\@';                                                                                                                                                  \n");
            htmlStr.Append("		text-align:center;                                                                                                                                                       \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:none;                                                                                                                                                       \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	.xl25222214                                                                                                                                                                  \n");
            htmlStr.Append("		{padding:0px;                                                                                                                                                            \n");
            htmlStr.Append("		mso-ignore:padding;                                                                                                                                                      \n");
            htmlStr.Append("		color:windowtext;                                                                                                                                                        \n");
            htmlStr.Append("		font-size:1.0pt;                                                                                                                                                         \n");
            htmlStr.Append("		font-weight:400;                                                                                                                                                         \n");
            htmlStr.Append("		font-style:normal;                                                                                                                                                       \n");
            htmlStr.Append("		text-decoration:none;                                                                                                                                                    \n");
            htmlStr.Append("		font-family:'Times New Roman', serif;                                                                                                                                    \n");
            htmlStr.Append("		mso-font-charset:0;                                                                                                                                                      \n");
            htmlStr.Append("		mso-number-format:General;                                                                                                                                               \n");
            htmlStr.Append("		text-align:general;                                                                                                                                                      \n");
            htmlStr.Append("		vertical-align:bottom;                                                                                                                                                   \n");
            htmlStr.Append("		border-top:.5pt solid windowtext;                                                                                                                                        \n");
            htmlStr.Append("		border-right:.5pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("		border-bottom:none;                                                                                                                                                      \n");
            htmlStr.Append("		border-left:none;                                                                                                                                                        \n");
            htmlStr.Append("		background:white;                                                                                                                                                        \n");
            htmlStr.Append("		mso-pattern:black none;                                                                                                                                                  \n");
            htmlStr.Append("		white-space:nowrap;}                                                                                                                                                     \n");
            htmlStr.Append("	-->                                                                                                                                                                          \n");
            htmlStr.Append("	</style>                                                                                                                                                                     \n"); htmlStr.Append("</head>                                                                                                                                                                          \n");
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


                htmlStr.Append("	<table border=0 cellpadding=0 cellspacing=0 width=737 class=xl6922214																											\n");
                htmlStr.Append("	 style='border-collapse:collapse;table-layout:fixed;width:660pt'>                                                                                                               \n");
                htmlStr.Append("	 <col class=xl6922214 width=12 style='mso-width-source:userset;mso-width-alt:                                                                                                   \n");
                htmlStr.Append("	 426;width:11.25pt'>                                                                                                                                                                \n");
                htmlStr.Append("	 <col class=xl6922214 width=34 style='mso-width-source:userset;mso-width-alt:                                                                                                   \n");
                htmlStr.Append("	 1194;width:43.75pt'>                                                                                                                                                              \n");
                htmlStr.Append("	 <col class=xl6922214 width=26 style='mso-width-source:userset;mso-width-alt:                                                                                                   \n");
                htmlStr.Append("	 938;width:31.25pt'>                                                                                                                                                               \n");
                htmlStr.Append("	 <col class=xl6922214 width=48 style='mso-width-source:userset;mso-width-alt:                                                                                                   \n");
                htmlStr.Append("	 1706;width:32.5pt'>                                                                                                                                                              \n");
                htmlStr.Append("	 <col class=xl6922214 width=4 style='mso-width-source:userset;mso-width-alt:                                                                                                    \n");
                htmlStr.Append("	 142;width:3.75pt'>                                                                                                                                                                \n");
                htmlStr.Append("	 <col class=xl6922214 width=41 style='mso-width-source:userset;mso-width-alt:                                                                                                   \n");
                htmlStr.Append("	 1450;width:26.25pt'>                                                                                                                                                              \n");
                htmlStr.Append("	 <col class=xl6922214 width=38 style='mso-width-source:userset;mso-width-alt:                                                                                                   \n");
                htmlStr.Append("	 1336;width:31.25pt'>                                                                                                                                                              \n");
                htmlStr.Append("	 <col class=xl6922214 width=14 style='mso-width-source:userset;mso-width-alt:                                                                                                   \n");
                htmlStr.Append("	 512;width:13.75pt'>                                                                                                                                                               \n");
                htmlStr.Append("	 <col class=xl6922214 width=23 style='mso-width-source:userset;mso-width-alt:                                                                                                   \n");
                htmlStr.Append("	 824;width:21.25pt'>                                                                                                                                                               \n");
                htmlStr.Append("	 <col class=xl6922214 width=13 style='mso-width-source:userset;mso-width-alt:                                                                                                   \n");
                htmlStr.Append("	 455;width:12.5pt'>                                                                                                                                                               \n");
                htmlStr.Append("	 <col class=xl6922214 width=25 style='mso-width-source:userset;mso-width-alt:                                                                                                   \n");
                htmlStr.Append("	 881;width:23.75pt'>                                                                                                                                                               \n");
                htmlStr.Append("	 <col class=xl6922214 width=9 span=2 style='mso-width-source:userset;                                                                                                           \n");
                htmlStr.Append("	 mso-width-alt:312;width:8.75pt'>                                                                                                                                                  \n");
                htmlStr.Append("	 <col class=xl6922214 width=19 style='mso-width-source:userset;mso-width-alt:                                                                                                   \n");
                htmlStr.Append("	 682;width:17.5pt'>                                                                                                                                                               \n");
                htmlStr.Append("	 <col class=xl6922214 width=0 style='display:none;mso-width-source:userset;                                                                                                     \n");
                htmlStr.Append("	 mso-width-alt:369'>                                                                                                                                                            \n");
                htmlStr.Append("	 <col class=xl6922214 width=26 span=2 style='mso-width-source:userset;                                                                                                          \n");
                htmlStr.Append("	 mso-width-alt:938;width:31.25pt'>                                                                                                                                                 \n");
                htmlStr.Append("	 <col class=xl6922214 width=25 style='mso-width-source:userset;mso-width-alt:                                                                                                   \n");
                htmlStr.Append("	 881;width:23.75pt'>                                                                                                                                                               \n");
                htmlStr.Append("	 <col class=xl6922214 width=18 style='mso-width-source:userset;mso-width-alt:                                                                                                   \n");
                htmlStr.Append("	 654;width:17.5pt'>                                                                                                                                                               \n");
                htmlStr.Append("	 <col class=xl6922214 width=29 style='mso-width-source:userset;mso-width-alt:                                                                                                   \n");
                htmlStr.Append("	 1024;width:27.5pt'>                                                                                                                                                              \n");
                htmlStr.Append("	 <col class=xl6922214 width=39 style='mso-width-source:userset;mso-width-alt:                                                                                                   \n");
                htmlStr.Append("	 1393;width:36.25pt'>                                                                                                                                                              \n");
                htmlStr.Append("	 <col class=xl6922214 width=33 style='mso-width-source:userset;mso-width-alt:                                                                                                   \n");
                htmlStr.Append("	 1166;width:31.25pt'>                                                                                                                                                              \n");
                htmlStr.Append("	 <col class=xl6922214 width=26 style='mso-width-source:userset;mso-width-alt:                                                                                                   \n");
                htmlStr.Append("	 938;width:31.25pt'>                                                                                                                                                               \n");
                htmlStr.Append("	 <col class=xl6922214 width=12 style='mso-width-source:userset;mso-width-alt:                                                                                                   \n");
                htmlStr.Append("	 426;width:11.25pt'>                                                                                                                                                                \n");
                htmlStr.Append("	 <col class=xl6922214 width=41 style='mso-width-source:userset;mso-width-alt:                                                                                                   \n");
                htmlStr.Append("	 1450;width:26.25pt'>                                                                                                                                                              \n");
                htmlStr.Append("	 <col class=xl6922214 width=34 style='mso-width-source:userset;mso-width-alt:                                                                                                   \n");
                htmlStr.Append("	 1194;width:31.25pt'>                                                                                                                                                              \n");
                htmlStr.Append("	 <col class=xl6922214 width=26 span=2 style='mso-width-source:userset;                                                                                                          \n");
                htmlStr.Append("	 mso-width-alt:938;width:31.25pt'>                                                                                                                                                 \n");
                htmlStr.Append("	 <col class=xl6922214 width=22 style='mso-width-source:userset;mso-width-alt:                                                                                                   \n");
                htmlStr.Append("	 768;width:20.0pt'>                                                                                                                                                               \n");
                htmlStr.Append("	 <col class=xl6922214 width=27 style='mso-width-source:userset;mso-width-alt:                                                                                                   \n");
                htmlStr.Append("	 967;width:31.25pt'>                                                                                                                                                               \n");
                htmlStr.Append("	 <col class=xl6922214 width=12 style='mso-width-source:userset;mso-width-alt:                                                                                                   \n");
                htmlStr.Append("	 426;width:11.25pt'>                                                                                                                                                                \n");
                htmlStr.Append("	                                                                                                                                                                                \n");
                htmlStr.Append("	 <tr height=33 style='mso-height-source:userset;height:31.32pt'>                                                                                                                \n");
                htmlStr.Append("	  <td height=33 class=xl7122214 style='height:31.32pt' align=left valign=top><![if !vml]><span style='mso-ignore:vglayout;                                                                      \n");
                htmlStr.Append("	  position:absolute;z-index:2;margin-left:2px;margin-top:17px;width:146px;                                                                                                      \n");
                htmlStr.Append("	  height:77px'><img width=146 height=77                                                                                                                                         \n");
                htmlStr.Append("	  src='D:/webproject/e-invoice-ws/02.Web/EInvoice/img/WEBCASH-LOGO_001.jpg' v:shapes='Picture_x0020_1'></span><![endif]><span                                                 \n");
                htmlStr.Append("	  style='mso-ignore:vglayout2'>                                                                                                                                                 \n");
                htmlStr.Append("	  <table cellpadding=0 cellspacing=0>                                                                                                                                           \n");
                htmlStr.Append("	   <tr>                                                                                                                                                                         \n");
                htmlStr.Append("	    <td height=33 width=12 style='height:31.32pt;width:11.25pt'>&nbsp;</td>                                                                                         \n");
                htmlStr.Append("	   </tr>                                                                                                                                                                        \n");
                htmlStr.Append("	  </table>                                                                                                                                                                      \n");
                htmlStr.Append("	  </span></td>                                                                                                                                                                  \n");
                htmlStr.Append("	  <td class=xl8522214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl8522214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl8622214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl8622214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl8722214 colspan=7>" + dt.Rows[0]["Seller_Name"] + "</td>                                                                                                                         \n");
                htmlStr.Append("	  <td class=xl8622214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl8622214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl8622214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl8622214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl8622214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl8622214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl8622214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl8622214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl8622214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl8622214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl8622214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl8622214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl8622214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl8622214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl8622214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl8622214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl8622214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl7222214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl7322214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	 </tr>                                                                                                                                                                          \n");
                htmlStr.Append("	 <tr class=xl7822214 height=36 style='mso-height-source:userset;height:29.75pt'>                                                                                                \n");
                htmlStr.Append("	  <td height=36 class=xl8022214 style='height:29.75pt'>&nbsp;</td>                                                                                                              \n");
                htmlStr.Append("	  <td class=xl7822214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl7822214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl7822214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl7822214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td colspan=26 class=xl17722214 width=613 style='border-right:.5pt solid black;                                                                                               \n");
                htmlStr.Append("	  width:463pt'>Mã s&#7889;                                                                                                                                                      \n");
                htmlStr.Append("	  thu&#7871; </font><font class='font622214'>(Tax Code)</font>:<span style='mso-spacerun:yes'>  </span><font class='font1522214'>" + dt.Rows[0]["Seller_TaxCode"] + "</font></font></td>     \n");
                htmlStr.Append("	 </tr>                                                                                                                                                                          \n");
                htmlStr.Append("  <tr class=xl7822214 height=25 style='mso-height-source:userset;height:19.05pt'>								\n");
                htmlStr.Append("    <td height=25 class=xl8022214 style='height:19.05pt'>&nbsp;</td>                                            \n");
                htmlStr.Append("    <td class=xl7822214>&nbsp;</td>                                                                             \n");
                htmlStr.Append("    <td class=xl7822214>&nbsp;</td>                                                                             \n");
                htmlStr.Append("    <td class=xl7822214>&nbsp;</td>                                                                             \n");
                htmlStr.Append("    <td class=xl7822214>&nbsp;</td>                                                                             \n");
                htmlStr.Append("    <td colspan=26 class=xl17922214 width=613 style='border-right:.5pt solid black;                             \n");
                htmlStr.Append("    width:463pt'>&#272;&#7883;a ch&#7881; <font class='font622214'>(Address)</font><font                        \n");
                htmlStr.Append("    class='font1022214'>: " + dt.Rows[0]["seller_address"] + "</td>                                                           \n");
                htmlStr.Append("   </tr>                                                                                                        \n");

                htmlStr.Append("	  <tr class=xl7822214 height=25 style='mso-height-source:userset;height:29.75pt'>                                                                                               \n");
                htmlStr.Append("	  <td height=25 class=xl8022214 style='height:29.75pt'>&nbsp;</td>                                                                                                              \n");
                htmlStr.Append("	  <td class=xl7822214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl7822214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl7822214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl7822214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td colspan=26 class=xl17922214 width=613 style='border-right:.5pt solid black;                                                                                               \n");
                htmlStr.Append("	  width:463pt'>&#272;i&#7879;n tho&#7841;i <font class='font622214'>(Tel)</font><font                                                                                           \n");
                htmlStr.Append("	  class='font1022214'>: " + dt.Rows[0]["Seller_Tel"] + " - Fax: " + dt.Rows[0]["Seller_Fax"] + "</td>                                                                                                     \n");
                htmlStr.Append("	 </tr>                                                                                                                                                                          \n");
                htmlStr.Append("	                                                                                                                                                                                \n");
                htmlStr.Append("	 <tr class=xl7822214 height=25 style='mso-height-source:userset;height:29.75pt'>                                                                                                \n");
                htmlStr.Append("	  <td height=25 class=xl12322214 style='height:29.75pt'>&nbsp;</td>                                                                                                             \n");
                htmlStr.Append("	  <td class=xl12422214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl12422214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl12422214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl12422214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl12522214 colspan=26 style='border-right:.5pt solid black'>S&#7889;                                                                                                \n");
                htmlStr.Append("	  tài kho&#7843;n<font class='font1322214'> </font><font class='font622214'>(Acc.                                                                                               \n");
                htmlStr.Append("	  code)</font><font class='font1022214'>: " + dt.Rows[0]["Seller_AccountNo"] + " " + dt.Rows[0]["BANK_NM31"] + "</font></td>                                                                                   \n");
                htmlStr.Append("	 </tr>                                                                                                                                                                          \n");
                htmlStr.Append("	 <tr height=4 style='mso-height-source:userset;height:3.0pt'>                                                                                                                   \n");
                htmlStr.Append("	  <td height=4 class=xl23522214 style='height:3.0pt'>&nbsp;</td>                                                                                                                \n");
                htmlStr.Append("	  <td class=xl23622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl23622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl23622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl23622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl23622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl23622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl23622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl23622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl23622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl23622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl23622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl23622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl23622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl23622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl23622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl23622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl23622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl23622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl23622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl23622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl23622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl23622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl23622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl23622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl23622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl23622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl23622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl23622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl23622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl23722214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	 </tr>                                                                                                                                                                          \n");
                htmlStr.Append("	 <tr height=26 style='mso-height-source:userset;height:24.38pt'>                                                                                                                 \n");
                htmlStr.Append("	  <td height=26 class=xl6822214 style='height:24.38pt'>&nbsp;</td>                                                                                                               \n");
                htmlStr.Append("	  <td class=xl6922214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl12222214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl12222214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl12222214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td colspan=20 class=xl20722214>PHI&#7870;U XU&#7844;T KHO KIÊM V&#7852;N                                                                                                     \n");
                htmlStr.Append("	  CHUY&#7874;N N&#7896;I B&#7896;</td>                                                                                                                                          \n");
                htmlStr.Append("	  <td class=xl12222214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl12222214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl12222214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl12222214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl12222214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl7422214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	 </tr>                                                                                                                                                                          \n");
                htmlStr.Append("	                                                                                                                                                                                \n");
                htmlStr.Append("	 <tr height=6 style='mso-height-source:userset;height:26.20pt'>                                                                                                                 \n");
                htmlStr.Append("	  <td height=6 class=xl6822214 style='height:26.20pt'>&nbsp;</td>                                                                                                               \n");
                htmlStr.Append("	  <td class=xl6422214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl6422214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl6422214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl6422214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td colspan=19  class=xl16422214>DELIVERY NOTE CUM INTERNAL                                                                                                                   \n");
                htmlStr.Append("	  TRANSPORT DOCUMENT</td>                                                                                                                                                       \n");
                htmlStr.Append("	  <td class=xl14622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl7522214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl7522214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl7522214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl7522214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl7522214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl7422214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	 </tr>                                                                                                                                                                          \n");
                htmlStr.Append("	 <tr height=20 style='mso-height-source:userset;height:18.75pt'>                                                                                                                 \n");
                htmlStr.Append("	  <td height=20 class=xl6822214 style='height:18.75pt'>&nbsp;</td>                                                                                                               \n");
                htmlStr.Append("	  <td class=xl6422214 colspan=21>&nbsp;</td>                                                                                                                                    \n");
                htmlStr.Append("	  <td class=xl15122214 colspan=4>Ký hi&#7879;u (<font class='font1822214'>Serial</font><font                                                                                    \n");
                htmlStr.Append("	  class='font2122214'>):</font></td>                                                                                                                                            \n");
                htmlStr.Append("	  <td class=xl8422214 colspan=4>&nbsp;&nbsp;" + dt.Rows[0]["templateCode"] + "" + dt.Rows[0]["InvoiceSerialNo"] + "</td>                                                                                                 \n");
                htmlStr.Append("	  <td class=xl7422214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	 </tr>                                                                                                                                                                          \n");
                htmlStr.Append("	 <tr height=20 style='mso-height-source:userset;height:18.75pt'>                                                                                                                 \n");
                htmlStr.Append("	  <td height=20 class=xl6822214 style='height:18.75pt'>&nbsp;</td>                                                                                                               \n");
                htmlStr.Append("	  <td colspan=21 class=xl16322214>&nbsp;</td>                                                                                                                                   \n");
                htmlStr.Append("	  <td class=xl12822214 colspan=4><font class='font1422214'>S&#7889; (</font><font                                                                                               \n");
                htmlStr.Append("	  class='font1822214'>No</font><font class='font1422214'>.):</font></td>                                                                                                        \n");
                htmlStr.Append("	  <td colspan=4 class=xl14522214>" + dt.Rows[0]["InvoiceNumber"] + "&nbsp;&nbsp;&nbsp;&nbsp;</td>                                                                                                                   \n");
                htmlStr.Append("	  <td class=xl7422214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	 </tr>                                                                                                                                                                          \n");
                htmlStr.Append("	 <tr height=20 style='mso-height-source:userset;height:18.75pt'>                                                                                                                 \n");
                htmlStr.Append("	  <td height=20 class=xl6822214 style='height:18.75pt'>&nbsp;</td>                                                                                                               \n");
                htmlStr.Append("	  <td class=xl6422214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl6422214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl6422214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl6422214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	  <td colspan=20 class=xl16222214>Mã CQT (<font class='font622214'>Veification                                                                                                  \n");
                htmlStr.Append("	  code</font><font class='font722214'>) : </font><font class='font2022214'>" + dt.Rows[0]["cqt_mccqt_id"] + "</font></td>                                                                                         \n");
                htmlStr.Append("	  <td class=xl12722214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl12822214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl14522214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl14522214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl14522214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl7422214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	 </tr>                                                                                                                                                                          \n");
                htmlStr.Append("	 <tr class=xl10622214 height=25 style='mso-height-source:userset;height:29.75pt'>                                                                                               \n");
                htmlStr.Append("	  <td height=25 class=xl10722214 style='height:29.75pt'>&nbsp;</td>                                                                                                             \n");
                htmlStr.Append("	  <td class=xl10622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl10622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl10622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl10622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl10622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td colspan=16 class=xl16222214>Ngày <font class='font522214'>(</font><font                                                                                                   \n");
                htmlStr.Append("	  class='font622214'>date</font><font class='font522214'>)</font><font                                                                                                          \n");
                htmlStr.Append("	  class='font722214'><span style='mso-spacerun:yes'>  " + dt.Rows[0]["invoiceissueddate_dd"] + "    </span>tháng </font><font                                                                                \n");
                htmlStr.Append("	  class='font522214'>(</font><font class='font622214'>month</font><font                                                                                                         \n");
                htmlStr.Append("	  class='font522214'>)</font><font class='font722214'><span                                                                                                                     \n");
                htmlStr.Append("	  style='mso-spacerun:yes'>  " + dt.Rows[0]["invoiceissueddate_mm"] + "   </span>n&#259;m </font><font class='font522214'>(</font><font                                                                      \n");
                htmlStr.Append("	  class='font622214'>year</font><font class='font522214'>)</font><font                                                                                                          \n");
                htmlStr.Append("	  class='font722214'><span style='mso-spacerun:yes'> " + dt.Rows[0]["invoiceissueddate_yyyy"] + " </span></font></td>                                                                                        \n");
                htmlStr.Append("	  <td class=xl10622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl10622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td colspan=6 class=xl16122214></td>                                                                                                                                          \n");
                htmlStr.Append("	  <td class=xl10822214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	 </tr>                                                                                                                                                                          \n");
                htmlStr.Append("	 <tr height=4 style='mso-height-source:userset;height:3.0pt'>                                                                                                                   \n");
                htmlStr.Append("	  <td height=4 class=xl24722214 style='height:3.0pt'>&nbsp;</td>                                                                                                                \n");
                htmlStr.Append("	  <td class=xl24822214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24822214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24822214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24822214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24822214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24822214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24822214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24822214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24822214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24822214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24822214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24922214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24822214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24822214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24922214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl25022214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl25022214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl25122214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl25122214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24822214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24822214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24822214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24822214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24822214 style='border-top:none'>&nbsp;</td>                                                                                                                      \n");
                htmlStr.Append("	  <td class=xl24822214 style='border-top:none'>&nbsp;</td>                                                                                                                      \n");
                htmlStr.Append("	  <td class=xl24822214 style='border-top:none'>&nbsp;</td>                                                                                                                      \n");
                htmlStr.Append("	  <td class=xl24822214 style='border-top:none'>&nbsp;</td>                                                                                                                      \n");
                htmlStr.Append("	  <td class=xl24822214 style='border-top:none'>&nbsp;</td>                                                                                                                      \n");
                htmlStr.Append("	  <td class=xl24822214 style='border-top:none'>&nbsp;</td>                                                                                                                      \n");
                htmlStr.Append("	  <td class=xl25222214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	 </tr>                                                                                                                                                                          \n");
                htmlStr.Append("	 <tr class=xl13122214 height=21 style='mso-height-source:userset;height:20.07pt'>                                                                                               \n");
                htmlStr.Append("	  <td height=21 class=xl12922214 style='height:20.07pt'>&nbsp;</td>                                                                                                             \n");
                htmlStr.Append("	  <td class=xl13022214 colspan=20>C&#259;n c&#7913; l&#7879;nh &#273;i&#7873;u                                                                                                  \n");
                htmlStr.Append("	  &#273;&#7897;ng s&#7889; (<font class='font622214'>Ordering no</font><font                                                                                                    \n");
                htmlStr.Append("	  class='font722214'>) : </font>" + dt.Rows[0]["attribute_01"] + "</td>                                                                                                                         \n");
                htmlStr.Append("	  <td class=xl13022214 colspan=9>Ngày ( <font class='font622214'>date</font><font                                                                                               \n");
                htmlStr.Append("	  class='font722214'> ) : </font> " + dt.Rows[0]["attribute_02"] + "</td>                                                                                                                       \n");
                htmlStr.Append("	  <td class=xl13222214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	 </tr>                                                                                                                                                                          \n");
                htmlStr.Append("	 <tr class=xl13122214 height=21 style='mso-height-source:userset;height:20.07pt'>                                                                                               \n");
                htmlStr.Append("	  <td height=21 class=xl12922214 style='height:20.07pt'>&nbsp;</td>                                                                                                             \n");
                htmlStr.Append("	  <td class=xl13022214 colspan=20>c&#7911;a ( <font class='font622214'>By </font><font                                                                                          \n");
                htmlStr.Append("	  class='font722214'>) :</font> " + dt.Rows[0]["attribute_03"] + "</td>                                                                                                                         \n");
                htmlStr.Append("	  <td class=xl13022214 colspan=9>v&#7873; vi&#7879;c ( <font class='font1822214'>for                                                                                            \n");
                htmlStr.Append("	  </font><font class='font1422214'>) :</font> " + dt.Rows[0]["attribute_04"] + "</td>                                                                                                           \n");
                htmlStr.Append("	  <td class=xl13222214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	 </tr>                                                                                                                                                                          \n");
                htmlStr.Append("	 <tr class=xl13122214 height=21 style='mso-height-source:userset;height:20.07pt'>                                                                                               \n");
                htmlStr.Append("	  <td height=21 class=xl12922214 style='height:20.07pt'>&nbsp;</td>                                                                                                             \n");
                htmlStr.Append("	  <td class=xl13022214 colspan=20>Tên ng&#432;&#7901;i v&#7853;n chuy&#7875;n (<font                                                                                            \n");
                htmlStr.Append("	  class='font622214'>Full name of transporter</font><font class='font722214'>)                                                                                                  \n");
                htmlStr.Append("	  :</font> " + dt.Rows[0]["attribute_05"] + "</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl130222141 colspan=9>H&#7907;p &#273;&#7891;ng s&#7889; (<font                                                                                                     \n");
                htmlStr.Append("	  class='font1822214'>Contract no</font><font class='font1422214'>) : </font>" + dt.Rows[0]["attribute_06"] + "</td>                                                                            \n");
                htmlStr.Append("	  <td class=xl13222214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	 </tr>                                                                                                                                                                          \n");
                htmlStr.Append("	 <tr class=xl13122214 height=21 style='mso-height-source:userset;height:20.07pt'>                                                                                               \n");
                htmlStr.Append("	  <td height=21 class=xl12922214 style='height:20.07pt'>&nbsp;</td>                                                                                                             \n");
                htmlStr.Append("	  <td class=xl13022214 colspan=29>Ph&#432;&#417;ng ti&#7879;n v&#7853;n                                                                                                         \n");
                htmlStr.Append("	  chuy&#7875;n (<font class='font1822214'>Transportation</font><font                                                                                                            \n");
                htmlStr.Append("	  class='font1422214'>) : </font> " + dt.Rows[0]["attribute_07"] + "</td>                                                                                                                       \n");
                htmlStr.Append("	  <td class=xl13222214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	 </tr>                                                                                                                                                                          \n");
                htmlStr.Append("	 <tr class=xl13122214 height=21 style='mso-height-source:userset;height:20.07pt'>                                                                                               \n");
                htmlStr.Append("	  <td height=21 class=xl12922214 style='height:20.07pt'>&nbsp;</td>                                                                                                             \n");
                htmlStr.Append("	  <td class=xl13022214 colspan=29>Xu&#7845;t t&#7841;i kho (<font                                                                                                               \n");
                htmlStr.Append("	  class='font1822214'>Stock out at</font><font class='font1422214'>) : " + dt.Rows[0]["attribute_08"] + "</font></td>                                                                           \n");
                htmlStr.Append("	  <td class=xl13222214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	 </tr>                                                                                                                                                                          \n");
                htmlStr.Append("	 <tr class=xl13122214 height=21 style='mso-height-source:userset;height:20.07pt'>                                                                                               \n");
                htmlStr.Append("	  <td height=21 class=xl12922214 style='height:20.07pt'>&nbsp;</td>                                                                                                             \n");
                htmlStr.Append("	  <td class=xl13022214 colspan=29>Nh&#7853;p t&#7841;i kho (<font                                                                                                               \n");
                htmlStr.Append("	  class='font1822214'>Stock in at</font><font class='font1422214'>) : " + dt.Rows[0]["attribute_09"] + "</font></td>                                                                            \n");
                htmlStr.Append("	  <td class=xl13222214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	 </tr>                                                                                                                                                                          \n");
                htmlStr.Append("	 <tr class=xl13122214 height=20 style='mso-height-source:userset;height:18.75pt'>                                                                                                \n");
                htmlStr.Append("	  <td height=20 class=xl12922214 style='height:18.75pt'>&nbsp;</td>                                                                                                              \n");
                htmlStr.Append("	  <td class=xl13022214 colspan=20>&#272;&#417;n v&#7883; ti&#7873;n t&#7879; (<font                                                                                             \n");
                htmlStr.Append("	  class='font622214'>Currency</font><font class='font722214'>):</font> " + dt.Rows[0]["CurrencyCodeUSD"] + "</td>                                                                                     \n");
                htmlStr.Append("	  <td class=xl13422214 colspan=9>T&#7927; giá (<font class='font622214'>Exchange                                                                                                \n");
                htmlStr.Append("	  rate</font><font class='font722214'>) : </font> " + dt.Rows[0]["tr_rate"] + "</td>                                                                                                         \n");
                htmlStr.Append("	  <td class=xl13222214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	 </tr>                                                                                                                                                                          \n");
                htmlStr.Append("	 <tr height=4 style='mso-height-source:userset;height:3.0pt'>                                                                                                                   \n");
                htmlStr.Append("	  <td height=4 class=xl23522214 style='height:3.0pt'>&nbsp;</td>                                                                                                                \n");
                htmlStr.Append("	  <td class=xl24522214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24522214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24522214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24522214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24522214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24522214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24522214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24522214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24522214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24522214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24522214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24522214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24522214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24522214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24522214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24522214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24522214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24522214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24522214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24522214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24522214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24522214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24522214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24522214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24522214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24522214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24522214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24522214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24522214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	  <td class=xl24622214>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("	 </tr>                                                                                                                                                                          \n");
                htmlStr.Append("	 <tr class=xl7822214 height=25 style='mso-height-source:userset;height:29.75pt'>                                                                                                \n");
                htmlStr.Append("	  <td height=25 class=xl7622214 style='height:29.75pt'>&nbsp;</td>                                                                                                              \n");
                htmlStr.Append("	  <td rowspan=2 class=xl18922214 width=34 style='border-bottom:.5pt solid black;                                                                                                \n");
                htmlStr.Append("	  border-top:none;width:31.25pt'>STT<br>                                                                                                                                           \n");
                htmlStr.Append("	    <font class='font1922214'>No</font></td>                                                                                                                                    \n");
                htmlStr.Append("	  <td colspan=12 rowspan=2 class=xl19122214 width=269 style='width:203pt'>Tên                                                                                                   \n");
                htmlStr.Append("	  hàng hóa, quy cách,<br>                                                                                                                                                       \n");
                htmlStr.Append("	    ph&#7849;m ch&#7845;t v&#7853;t t&#432; <br>                                                                                                                                \n");
                htmlStr.Append("	    (S&#7843;n ph&#7849;m hàng hóa)<span style='mso-spacerun:yes'>                                                                                                              \n");
                htmlStr.Append("	  </span><br>                                                                                                                                                                   \n");
                htmlStr.Append("	    <font class='font622214'>(Product name, specification)</font><font                                                                                                          \n");
                htmlStr.Append("	  class='font822214'></font></td>                                                                                                                                               \n");
                htmlStr.Append("	                                                                                                                                                                                \n");
                htmlStr.Append("	  <td colspan=3 rowspan=2 class=xl16622214 width=52 style='border-right:.5pt solid black;                                                                                       \n");
                htmlStr.Append("	  border-bottom:.5pt solid black;width:40pt'>Mã s&#7889;<br>                                                                                                                    \n");
                htmlStr.Append("	    <font class='font622214'>(Part no)</font></td>                                                                                                                              \n");
                htmlStr.Append("	  <td colspan=2 rowspan=2 class=xl16522214 width=43 style='border-right:.5pt solid black;                                                                                       \n");
                htmlStr.Append("	  border-bottom:.5pt solid black;width:33pt'>&#272;&#417;n v&#7883; tính<br>                                                                                                    \n");
                htmlStr.Append("	    <font class='font622214'>(Unit)</font><font class='font822214'></font></td>                                                                                                 \n");
                htmlStr.Append("	  <td colspan=4 class=xl17122214 width=127 style='border-right:.5pt solid black;                                                                                                \n");
                htmlStr.Append("	  border-left:none;width:96pt'>S&#7889; l&#432;&#7907;ng<span                                                                                                                   \n");
                htmlStr.Append("	  style='mso-spacerun:yes'> </span></td>                                                                                                                                        \n");
                htmlStr.Append("	  <td colspan=3 rowspan=2 class=xl16522214 width=87 style='border-right:.5pt solid black;                                                                                       \n");
                htmlStr.Append("	  border-bottom:.5pt solid black;width:65pt'>&#272;&#417;n giá<br>                                                                                                              \n");
                htmlStr.Append("	    <font class='font622214'>(Unit price)</font></td>                                                                                                                           \n");
                htmlStr.Append("	  <td colspan=4 rowspan=2 class=xl16522214 width=101 style='border-right:.5pt solid black;                                                                                      \n");
                htmlStr.Append("	  border-bottom:.5pt solid black;width:76pt'>Thành ti&#7873;n<br>                                                                                                               \n");
                htmlStr.Append("	    <font class='font622214'>(Amount)</font></td>                                                                                                                               \n");
                htmlStr.Append("	  <td class=xl7722214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	 </tr>                                                                                                                                                                          \n");
                htmlStr.Append("	 <tr class=xl7822214 height=48 style='mso-height-source:userset;height:45.00pt'>                                                                                                 \n");
                htmlStr.Append("	  <td height=48 class=xl7622214 style='height:45.00pt'>&nbsp;</td>                                                                                                               \n");
                htmlStr.Append("	                                                                                                                                                                                \n");
                htmlStr.Append("	  <td colspan=2 class=xl19122214 width=68 style='border-left:none;width:51pt'>Th&#7921;c                                                                                        \n");
                htmlStr.Append("	  xu&#7845;t<br>                                                                                                                                                                \n");
                htmlStr.Append("	    <font class='font622214'>(Out)</font></td>                                                                                                                                  \n");
                htmlStr.Append("	  <td colspan=2 class=xl19122214 width=59 style='border-left:none;width:45pt'>Th&#7921;c                                                                                        \n");
                htmlStr.Append("	  nh&#7853;p<br>                                                                                                                                                                \n");
                htmlStr.Append("	    <font class='font622214'>(In)</font></td>                                                                                                                                   \n");
                htmlStr.Append("	  <td class=xl7722214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	 </tr>                                                                                                                                                                          \n");
                htmlStr.Append("	 <tr class=xl7822214 height=19 style='mso-height-source:userset;height:18.20pt'>                                                                                                \n");
                htmlStr.Append("	  <td height=19 class=xl7622214 style='height:18.20pt'>&nbsp;</td>                                                                                                              \n");
                htmlStr.Append("	  <td class=xl9822214 width=34 style='border-left:none;width:31.25pt'>A</td>                                                                                                       \n");
                htmlStr.Append("	  <td colspan=12 class=xl19122214 width=269 style='border-left:none;width:203pt'>B</td>                                                                                         \n");
                //htmlStr.Append("	  <td class=xl9922214 width=0>&nbsp;</td>                                                                                                                                       \n");
                htmlStr.Append("	  <td colspan=3 class=xl17222214 width=52 style='border-right:.5pt solid black;                                                                                                 \n");
                htmlStr.Append("	  width:40pt'>C</td>                                                                                                                                                            \n");
                htmlStr.Append("	  <td colspan=2 class=xl17122214 width=43 style='border-right:.5pt solid black;                                                                                                 \n");
                htmlStr.Append("	  border-left:none;width:33pt'>D</td>                                                                                                                                           \n");
                htmlStr.Append("	  <td colspan=2 class=xl17122214 width=68 style='border-right:.5pt solid black;                                                                                                 \n");
                htmlStr.Append("	  border-left:none;width:51pt'>1</td>                                                                                                                                           \n");
                htmlStr.Append("	  <td colspan=2 class=xl17122214 width=59 style='border-right:.5pt solid black;                                                                                                 \n");
                htmlStr.Append("	  border-left:none;width:45pt'>&nbsp;</td>                                                                                                                                      \n");
                htmlStr.Append("	  <td colspan=3 class=xl17122214 width=87 style='border-right:.5pt solid black;                                                                                                 \n");
                htmlStr.Append("	  border-left:none;width:65pt'>2</td>                                                                                                                                           \n");
                htmlStr.Append("	  <td colspan=4 class=xl17122214 width=101 style='border-right:.5pt solid black;                                                                                                \n");
                htmlStr.Append("	  border-left:none;width:76pt'>3 = 1 x 2</td>                                                                                                                                   \n");
                htmlStr.Append("	  <td class=xl7722214>&nbsp;</td>                                                                                                                                               \n");
                htmlStr.Append("	 </tr>                                                                                                                                                                          \n");
                htmlStr.Append("	                                                                                                                                                                                \n");

                v_rowHeight = "50.0pt"; //"26.5pt";
                v_rowHeightEmpty = "50.0pt";
                v_rowHeightNumber = 50;

                v_rowHeightLast = "50.0pt";// "23.5pt";
                v_rowHeightLastNumber = 50;// 23.5;
                v_rowHeightEmptyLast = "50.pt"; //"23.5pt";


                for (int dtR = 0; dtR < page[k]; dtR++)  //page[k]
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

                        htmlStr.Append("		 <tr class=xl7822214 height=29 style='mso-height-source:userset;height:27.56pt'>															\n");
                        htmlStr.Append("		  <td height=29 class=xl7622214 style='height:27.56pt'>&nbsp;</td>                                                                          \n");
                        htmlStr.Append("		  <td class=xl10922214 width=34 style='border-left:none;width:31.25pt'>&nbsp;" + dt_d.Rows[dtR][7] + "</td>                                                \n");
                        htmlStr.Append("		  <td colspan=12 class=xl11422214 width=269 style='border-right:.5pt solid black;                                                           \n");
                        htmlStr.Append("		  border-left:none;width:203pt'>&nbsp;" + dt_d.Rows[dtR][0] + "</td>                                                                                    \n");
                        htmlStr.Append("		  <td colspan=3 class=xl18422214 width=52 style='border-right:.5pt solid black;border-top:.5pt solid black;                                 \n");
                        htmlStr.Append("		  width:40pt'>&nbsp;</td>                                                                                                                   \n");
                        htmlStr.Append("		  <td colspan=2 class=xl19622214 width=43 style='border-right:.5pt solid black;                                                             \n");
                        htmlStr.Append("		  border-left:none;width:33pt'>&nbsp;" + dt_d.Rows[dtR][1] + "</td>                                                                                     \n");
                        htmlStr.Append("		  <td colspan=2 class=xl19822214 width=68 style='border-right:.5pt solid black;                                                             \n");
                        htmlStr.Append("		  border-left:none;width:51pt'>" + dt_d.Rows[dtR][2] + "&nbsp;</td>                                                                                     \n");
                        htmlStr.Append("		  <td colspan=2 class=xl19822214 width=59 style='border-right:.5pt solid black;                                                             \n");
                        htmlStr.Append("		  border-left:none;width:45pt'>&nbsp;</td>                                                                                                  \n");
                        htmlStr.Append("		  <td colspan=3 class=xl17422214 width=87 style='border-right:.5pt solid black;                                                             \n");
                        htmlStr.Append("		  border-left:none;width:65pt'>" + dt_d.Rows[dtR][3] + "&nbsp;</td>                                                                                     \n");
                        htmlStr.Append("		  <td colspan=4 class=xl20022214 width=101 style='border-left:none;width:76pt'>" + dt_d.Rows[dtR][4] + "&nbsp;</td>                                     \n");
                        htmlStr.Append("		  <td class=xl7722214>&nbsp;</td>                                                                                                           \n");
                        htmlStr.Append("		 </tr>                                                                                                                                      \n");

                    }
                    else if (dtR == page[k] - 1)//dong cuoi moi trang
                    {
                        if (k < v_countNumberOfPages - 1) //trang giua
                        {

                            htmlStr.Append("			<tr class=xl7822214 height=29 style='mso-height-source:userset;height:27.56pt'>													\n");
                            htmlStr.Append("			  <td height=29 class=xl7622214 style='height:27.56pt'>&nbsp;</td>                                                              \n");
                            htmlStr.Append("			  <td class=xl11022214 width=34 style='border-left:none;width:31.25pt'>&nbsp;" + dt_d.Rows[dtR][7] + "</td>                                    \n");
                            htmlStr.Append("			  <td colspan=12 class=xl22422214 width=269 style='border-right:.5pt solid black;                                               \n");
                            htmlStr.Append("			  border-left:none;width:203pt'>&nbsp;" + dt_d.Rows[dtR][0] + "</td>                                                                        \n");
                            htmlStr.Append("			  <td colspan=2 class=xl22622214 width=52 style='border-right:.5pt solid black;                                                 \n");
                            htmlStr.Append("			  width:40pt'>&nbsp;</td>                                                                                                       \n");
                            htmlStr.Append("			  <td colspan=2 class=xl22822214 width=43 style='border-right:.5pt solid black;                                                 \n");
                            htmlStr.Append("			  border-left:none;width:33pt'>" + dt_d.Rows[dtR][1] + "</td>                                                                               \n");
                            htmlStr.Append("			  <td colspan=2 class=xl22922214 width=68 style='border-right:.5pt solid black;                                                 \n");
                            htmlStr.Append("			  border-left:none;width:51pt'>" + dt_d.Rows[dtR][2] + "&nbsp;</td>                                                                         \n");
                            htmlStr.Append("			  <td colspan=2 class=xl22922214 width=59 style='border-right:.5pt solid black;                                                 \n");
                            htmlStr.Append("			  border-left:none;width:45pt'>&nbsp;</td>                                                                                      \n");
                            htmlStr.Append("			  <td colspan=3 class=xl15522214 width=87 style='border-right:.5pt solid black;                                                 \n");
                            htmlStr.Append("			  border-left:none;width:65pt'>" + dt_d.Rows[dtR][3] + "&nbsp;</td>                                                                         \n");
                            htmlStr.Append("			  <td colspan=4 class=xl21922214 width=101 style='border-left:none;width:76pt'>" + dt_d.Rows[dtR][4] + "&nbsp;</td>                         \n");
                            htmlStr.Append("			  <td class=xl7722214>&nbsp;</td>                                                                                               \n");
                            htmlStr.Append("			 </tr>                                                                                                                          \n");
                        }
                        else // trang cuoi
                        {
                            if (dtR == rowsPerPage - 1) // du 11 dong
                            {

                                htmlStr.Append("			<tr class=xl7822214 height=29 style='mso-height-source:userset;height:27.56pt'>													\n");
                                htmlStr.Append("			  <td height=29 class=xl7622214 style='height:27.56pt'>&nbsp;</td>                                                              \n");
                                htmlStr.Append("			  <td class=xl11022214 width=34 style='border-left:none;width:31.25pt'>&nbsp;" + dt_d.Rows[dtR][7] + "</td>                                    \n");
                                htmlStr.Append("			  <td colspan=12 class=xl22422214 width=269 style='border-right:.5pt solid black;                                               \n");
                                htmlStr.Append("			  border-left:none;width:203pt'>&nbsp;" + dt_d.Rows[dtR][0] + "</td>                                                                        \n");
                                htmlStr.Append("			  <td colspan=2 class=xl22622214 width=52 style='border-right:.5pt solid black;                                                 \n");
                                htmlStr.Append("			  width:40pt'>&nbsp;</td>                                                                                                       \n");
                                htmlStr.Append("			  <td colspan=2 class=xl22822214 width=43 style='border-right:.5pt solid black;                                                 \n");
                                htmlStr.Append("			  border-left:none;width:33pt'>" + dt_d.Rows[dtR][1] + "</td>                                                                               \n");
                                htmlStr.Append("			  <td colspan=2 class=xl22922214 width=68 style='border-right:.5pt solid black;                                                 \n");
                                htmlStr.Append("			  border-left:none;width:51pt'>" + dt_d.Rows[dtR][2] + "&nbsp;</td>                                                                         \n");
                                htmlStr.Append("			  <td colspan=2 class=xl22922214 width=59 style='border-right:.5pt solid black;                                                 \n");
                                htmlStr.Append("			  border-left:none;width:45pt'>&nbsp;</td>                                                                                      \n");
                                htmlStr.Append("			  <td colspan=3 class=xl15522214 width=87 style='border-right:.5pt solid black;                                                 \n");
                                htmlStr.Append("			  border-left:none;width:65pt'>" + dt_d.Rows[dtR][3] + "&nbsp;</td>                                                                         \n");
                                htmlStr.Append("			  <td colspan=4 class=xl21922214 width=101 style='border-left:none;width:76pt'>" + dt_d.Rows[dtR][4] + "&nbsp;</td>                         \n");
                                htmlStr.Append("			  <td class=xl7722214>&nbsp;</td>                                                                                               \n");
                                htmlStr.Append("			 </tr>                                                                                                                          \n");

                            }
                            else
                            {

                                htmlStr.Append("			<tr class=xl7822214 height=29 style='mso-height-source:userset;height:27.56pt'>													\n");
                                htmlStr.Append("			  <td height=29 class=xl7622214 style='height:27.56pt'>&nbsp;</td>                                                              \n");
                                htmlStr.Append("			  <td class=xl11022214 width=34 style='border-left:none;width:31.25pt'>&nbsp;" + dt_d.Rows[dtR][7] + "</td>                                    \n");
                                htmlStr.Append("			  <td colspan=12 class=xl22422214 width=269 style='border-right:.5pt solid black;                                               \n");
                                htmlStr.Append("			  border-left:none;width:203pt'>&nbsp;" + dt_d.Rows[dtR][0] + "</td>                                                                        \n");
                                htmlStr.Append("			  <td colspan=2 class=xl22622214 width=52 style='border-right:.5pt solid black;                                                 \n");
                                htmlStr.Append("			  width:40pt'>&nbsp;</td>                                                                                                       \n");
                                htmlStr.Append("			  <td colspan=2 class=xl22822214 width=43 style='border-right:.5pt solid black;                                                 \n");
                                htmlStr.Append("			  border-left:none;width:33pt'>" + dt_d.Rows[dtR][1] + "</td>                                                                               \n");
                                htmlStr.Append("			  <td colspan=2 class=xl22922214 width=68 style='border-right:.5pt solid black;                                                 \n");
                                htmlStr.Append("			  border-left:none;width:51pt'>" + dt_d.Rows[dtR][2] + "&nbsp;</td>                                                                         \n");
                                htmlStr.Append("			  <td colspan=2 class=xl22922214 width=59 style='border-right:.5pt solid black;                                                 \n");
                                htmlStr.Append("			  border-left:none;width:45pt'>&nbsp;</td>                                                                                      \n");
                                htmlStr.Append("			  <td colspan=3 class=xl15522214 width=87 style='border-right:.5pt solid black;                                                 \n");
                                htmlStr.Append("			  border-left:none;width:65pt'>" + dt_d.Rows[dtR][3] + "&nbsp;</td>                                                                         \n");
                                htmlStr.Append("			  <td colspan=4 class=xl21922214 width=101 style='border-left:none;width:76pt'>" + dt_d.Rows[dtR][4] + "&nbsp;</td>                         \n");
                                htmlStr.Append("			  <td class=xl7722214>&nbsp;</td>                                                                                               \n");
                                htmlStr.Append("			 </tr>                                                                                                                          \n");
                            }

                        }
                    }
                    else
                    { // dong giua                                                                                                                                    

                        htmlStr.Append(" 			  <tr class=xl7822214 height=29 style='mso-height-source:userset;height:27.56pt'>														 \n");
                        htmlStr.Append(" 					  <td height=29 class=xl7622214 style='height:27.56pt'>&nbsp;</td>                                                             \n");
                        htmlStr.Append(" 					  <td class=xl11122214 width=34 style='border-left:none; width:31.25pt'>&nbsp;" + dt_d.Rows[dtR][7] + "</td>                                  \n");
                        htmlStr.Append(" 					  <td colspan=12 class=xl11422214 width=269 style='border-right:.5pt solid black;                                              \n");
                        htmlStr.Append(" 					  border-left:none;width:203pt'>&nbsp;" + dt_d.Rows[dtR][0] + "</td>                                                                       \n");
                        htmlStr.Append(" 					  <td colspan=3 class=xl18422214 width=52 style='border-right:.5pt solid black;                                                \n");
                        htmlStr.Append(" 					  width:40pt'>&nbsp;</td>                                                                                                      \n");
                        htmlStr.Append(" 					  <td colspan=2 class=xl18622214 width=43 style='border-right:.5pt solid black;                                                \n");
                        htmlStr.Append(" 					  border-left:none;width:33pt'>&nbsp;" + dt_d.Rows[dtR][1] + "</td>                                                                        \n");
                        htmlStr.Append(" 					  <td colspan=2 class=xl18722214 width=68 style='border-right:.5pt solid black;                                                \n");
                        htmlStr.Append(" 					  border-left:none;width:51pt'>&nbsp;" + dt_d.Rows[dtR][2] + "&nbsp;</td>                                                                  \n");
                        htmlStr.Append(" 					  <td colspan=2 class=xl18722214 width=59 style='border-right:.5pt solid black;                                                \n");
                        htmlStr.Append(" 					  border-left:none;width:45pt'>&nbsp;</td>                                                                                     \n");
                        htmlStr.Append(" 					  <td colspan=3 class=xl15222214 width=87 style='border-right:.5pt solid black;                                                \n");
                        htmlStr.Append(" 					  border-left:none;width:65pt'>&nbsp;" + dt_d.Rows[dtR][3] + "&nbsp;</td>                                                                  \n");
                        htmlStr.Append(" 					  <td colspan=4 class=xl18322214 width=101 style='border-left:none;width:76pt'>" + dt_d.Rows[dtR][4] + "&nbsp;</td>                        \n");
                        htmlStr.Append(" 					  <td class=xl7722214>&nbsp;</td>                                                                                              \n");
                        htmlStr.Append(" 					 </tr>                                                                                                                         \n");
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

                            htmlStr.Append(" 	<tr class=xl7822214 height=29 style='mso-height-source:userset;height:27.56pt'>												  \n");
                            htmlStr.Append(" 	  <td height=29 class=xl7622214 style='height:27.56pt'>&nbsp;</td>                                                            \n");
                            htmlStr.Append(" 	  <td class=xl11022214 width=34 style='border-left:none;width:31.25pt'>&nbsp;</td>                                               \n");
                            htmlStr.Append(" 	  <td colspan=12 class=xl22422214 width=269 style='border-right:.5pt solid black;                                             \n");
                            htmlStr.Append(" 	  border-left:none;width:203pt'>&nbsp;</td>                                                                                   \n");
                            htmlStr.Append(" 	  <td class=xl11322214 width=0>&nbsp;</td>                                                                                    \n");
                            htmlStr.Append(" 	  <td colspan=2 class=xl22622214 width=52 style='border-right:.5pt solid black;                                               \n");
                            htmlStr.Append(" 	  width:40pt'>&nbsp;</td>                                                                                                     \n");
                            htmlStr.Append(" 	  <td colspan=2 class=xl22822214 width=43 style='border-right:.5pt solid black;                                               \n");
                            htmlStr.Append(" 	  border-left:none;width:33pt'>&nbsp;</td>                                                                                    \n");
                            htmlStr.Append(" 	  <td colspan=2 class=xl22922214 width=68 style='border-right:.5pt solid black;                                               \n");
                            htmlStr.Append(" 	  border-left:none;width:51pt'>&nbsp;</td>                                                                                    \n");
                            htmlStr.Append(" 	  <td colspan=2 class=xl22922214 width=59 style='border-right:.5pt solid black;                                               \n");
                            htmlStr.Append(" 	  border-left:none;width:45pt'>&nbsp;</td>                                                                                    \n");
                            htmlStr.Append(" 	  <td colspan=3 class=xl15522214 width=87 style='border-right:.5pt solid black;                                               \n");
                            htmlStr.Append(" 	  border-left:none;width:65pt'>&nbsp;</td>                                                                                    \n");
                            htmlStr.Append(" 	  <td colspan=4 class=xl21922214 width=101 style='border-left:none;width:76pt'>&nbsp;</td>                                    \n");
                            htmlStr.Append(" 	  <td class=xl7722214>&nbsp;</td>                                                                                             \n");
                            htmlStr.Append(" 	 </tr>                                                                                                                        \n");

                        }
                        else
                        {
                            htmlStr.Append(" 	<tr class=xl7822214 height=29 style='mso-height-source:userset;height:27.56pt'>											  \n");
                            htmlStr.Append(" 	  <td height=29 class=xl7622214 style='height:27.56pt'>&nbsp;</td>                                                        \n");
                            htmlStr.Append(" 	  <td class=xl11122214 width=34 style='border-left:none;width:31.25pt'>&nbsp;</td>                                           \n");
                            htmlStr.Append(" 	  <td colspan=12 class=xl11422214 width=269 style='border-right:.5pt solid black;                                         \n");
                            htmlStr.Append(" 	  border-left:none;width:203pt'>&nbsp;</td>                                                                               \n");
                            htmlStr.Append(" 	  <td class=xl11522214 width=0>&nbsp;</td>                                                                                \n");
                            htmlStr.Append(" 	  <td colspan=2 class=xl18422214 width=52 style='border-right:.5pt solid black;                                           \n");
                            htmlStr.Append(" 	  width:40pt'>&nbsp;</td>                                                                                                 \n");
                            htmlStr.Append(" 	  <td colspan=2 class=xl18622214 width=43 style='border-right:.5pt solid black;                                           \n");
                            htmlStr.Append(" 	  border-left:none;width:33pt'>&nbsp;</td>                                                                                \n");
                            htmlStr.Append(" 	  <td colspan=2 class=xl18722214 width=68 style='border-right:.5pt solid black;                                           \n");
                            htmlStr.Append(" 	  border-left:none;width:51pt'>&nbsp;</td>                                                                                \n");
                            htmlStr.Append(" 	  <td colspan=2 class=xl18722214 width=59 style='border-right:.5pt solid black;                                           \n");
                            htmlStr.Append(" 	  border-left:none;width:45pt'>&nbsp;</td>                                                                                \n");
                            htmlStr.Append(" 	  <td colspan=3 class=xl15222214 width=87 style='border-right:.5pt solid black;                                           \n");
                            htmlStr.Append(" 	  border-left:none;width:65pt'>&nbsp;</td>                                                                                \n");
                            htmlStr.Append(" 	  <td colspan=4 class=xl18322214 width=101 style='border-left:none;width:76pt'>&nbsp;</td>                                \n");
                            htmlStr.Append(" 	  <td class=xl7722214>&nbsp;</td>                                                                                         \n");
                            htmlStr.Append(" 	 </tr>                                                                                                                    \n");
                        }
                    } // for

                }//Trang cuoi 11 dong

                if (k < v_countNumberOfPages - 1)
                {

                    htmlStr.Append(" 	<tr class=xl7822214 height=29 style='mso-height-source:userset;height:27.56pt'>												  \n");
                    htmlStr.Append(" 	  <td height=29 class=xl7622214 style='height:27.56pt'>&nbsp;</td>                                                            \n");
                    htmlStr.Append(" 	  <td class=xl11022214 width=34 style='border-left:none;width:31.25pt'>&nbsp;</td>                                               \n");
                    htmlStr.Append(" 	  <td colspan=12 class=xl22422214 width=269 style='border-right:.5pt solid black;                                             \n");
                    htmlStr.Append(" 	  border-left:none;width:203pt'>&nbsp;</td>                                                                                   \n");
                    htmlStr.Append(" 	  <td class=xl11322214 width=0>&nbsp;</td>                                                                                    \n");
                    htmlStr.Append(" 	  <td colspan=2 class=xl22622214 width=52 style='border-right:.5pt solid black;                                               \n");
                    htmlStr.Append(" 	  width:40pt'>&nbsp;</td>                                                                                                     \n");
                    htmlStr.Append(" 	  <td colspan=2 class=xl22822214 width=43 style='border-right:.5pt solid black;                                               \n");
                    htmlStr.Append(" 	  border-left:none;width:33pt'>&nbsp;</td>                                                                                    \n");
                    htmlStr.Append(" 	  <td colspan=2 class=xl22922214 width=68 style='border-right:.5pt solid black;                                               \n");
                    htmlStr.Append(" 	  border-left:none;width:51pt'>&nbsp;</td>                                                                                    \n");
                    htmlStr.Append(" 	  <td colspan=2 class=xl22922214 width=59 style='border-right:.5pt solid black;                                               \n");
                    htmlStr.Append(" 	  border-left:none;width:45pt'>&nbsp;</td>                                                                                    \n");
                    htmlStr.Append(" 	  <td colspan=3 class=xl15522214 width=87 style='border-right:.5pt solid black;                                               \n");
                    htmlStr.Append(" 	  border-left:none;width:65pt'>&nbsp;</td>                                                                                    \n");
                    htmlStr.Append(" 	  <td colspan=4 class=xl21922214 width=101 style='border-left:none;width:76pt'>&nbsp;</td>                                    \n");
                    htmlStr.Append(" 	  <td class=xl7722214>&nbsp;</td>                                                                                             \n");
                    htmlStr.Append(" 	 </tr>                                                                                                                        \n");



                    htmlStr.Append("	<table  border=0>                                                                                                                                                                                                 \n");
                    htmlStr.Append("		<tr height=5  style='height: 55pt'>                                                                                                                                                                \n");
                    htmlStr.Append("		</tr>      																																														\n");
                    htmlStr.Append("	</table>             																																										\n");



                }


            }// for k                                                                                                                             
            htmlStr.Append("  <tr class=xl14222214 height=29 style='mso-height-source:userset;height:27.56pt'>																				\n");
            htmlStr.Append("   <td height=29 class=xl13722214 style='height:27.56pt'>&nbsp;</td>                                                                                            \n");
            htmlStr.Append("   <td class=xl13822214 style='border-left:none'>&nbsp;</td>                                                                                                    \n");
            htmlStr.Append("   <td colspan=12 class=xl22022214>T&#7893;ng c&#7897;ng <font                                                                                                  \n");
            htmlStr.Append("   class='font1922214'>(Total)</font><font class='font822214'>:</font></td>                                                                                     \n");
            htmlStr.Append("   <td class=xl13922214>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("   <td class=xl13922214>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("   <td class=xl14022214 style='border-top:none'>&nbsp;</td>                                                                                                     \n");
            htmlStr.Append("   <td class=xl14022214 style='border-top:none'>&nbsp;</td>                                                                                                     \n");
            htmlStr.Append("   <td class=xl14022214 style='border-top:none'>&nbsp;</td>                                                                                                     \n");
            htmlStr.Append("   <td colspan=2 class=xl22122214>&nbsp;</td>                                                                                                                   \n");
            htmlStr.Append("   <td colspan=2 class=xl22222214 style='border-left:none'>&nbsp;</td>                                                                                          \n");
            htmlStr.Append("   <td colspan=3 class=xl15822214 style='border-right:.5pt solid black'>&nbsp;</td>                                                                             \n");
            htmlStr.Append("   <td colspan=4 class=xl22322214 style='border-left:none'>" + dt.Rows[0]["TOT_AMT_BK_64"] + "&nbsp;</td>                                                                            \n");
            htmlStr.Append("   <td class=xl14122214>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("  </tr>                                                                                                                                                         \n");
            htmlStr.Append("  <tr class=xl7822214 height=4 style='mso-height-source:userset;height:3.0pt'>                                                                                  \n");
            htmlStr.Append("   <td height=4 class=xl24122214 style='height:3.0pt'>&nbsp;</td>                                                                                               \n");
            htmlStr.Append("   <td colspan=12 class=xl24222214>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("   <td colspan=17 class=xl24322214 width=429 style='width:324pt'>&nbsp;</td>                                                                                    \n");
            htmlStr.Append("   <td class=xl24422214>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("  </tr>                                                                                                                                                         \n");
            htmlStr.Append("  <tr class=xl6922214 height=6 style='mso-height-source:userset;height:4.5pt'>                                                                                  \n");
            htmlStr.Append("   <td colspan=31 height=6 class=xl23822214 style='border-right:.5pt solid black;                                                                               \n");
            htmlStr.Append("   height:4.5pt'>&nbsp;</td>                                                                                                                                    \n");
            htmlStr.Append("  </tr>                                                                                                                                                         \n");
            htmlStr.Append("  <tr class=xl6922214 height=9 style='mso-height-source:userset;height:6.75pt'>                                                                                 \n");
            htmlStr.Append("   <td height=9 class=xl7122214 style='height:6.75pt;border-top:none'>&nbsp;</td>                                                                               \n");
            htmlStr.Append("   <td class=xl7222214  colspan=29 style='border-top:none'>&nbsp;</td>                                                                                          \n");
            htmlStr.Append("   <td class=xl7422214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("  </tr>                                                                                                                                                         \n");
            htmlStr.Append("  <tr class=xl14922214 height=26 style='mso-height-source:userset;height:19.95pt'>                                                                              \n");
            htmlStr.Append("   <td height=26 class=xl14722214 style='height:19.95pt'>&nbsp;</td>                                                                                            \n");
            htmlStr.Append("   <td colspan=8 class=xl9922214 width=228 style='width:171pt'>Ng&#432;&#7901;i                                                                                 \n");
            htmlStr.Append("   nh&#7853;n hàng <font class='font622214'>(Consignee)</font></td>                                                                                                                                         \n");
            htmlStr.Append("   <td class=xl14822214>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("   <td class=xl14822214>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("   <td class=xl14822214>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("   <td class=xl14822214>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("   <td class=xl14822214>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("   <td class=xl14822214>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("   <td class=xl14922214>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("   <td class=xl14922214>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("   <td class=xl14822214>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("   <td class=xl14822214>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("   <td colspan=11 class=xl9922214 width=315 style='width:237pt'>Ng&#432;&#7901;i                                                                                \n");
            htmlStr.Append("   xu&#7845;t hàng <font class='font622214'>(Consignor)</font></td>                                                                                                                                         \n");
            htmlStr.Append("   <td class=xl15022214>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("  </tr>                                                                                                                                                         \n");
            htmlStr.Append("  <tr class=xl9222214 height=18 style='height:16.5pt'>                                                                                                          \n");
            htmlStr.Append("   <td height=18 class=xl8922214 style='height:16.5pt'>&nbsp;</td>                                                                                              \n");
            htmlStr.Append("   <td colspan=8 class=xl20822214>(Ch&#7919; ký s&#7889; (n&#7871;u có))</td>                                                                                   \n");
            htmlStr.Append("   <td class=xl9222214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl9222214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl9222214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl9222214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl9222214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl9422214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl9222214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl9222214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl9222214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl9222214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td colspan=11 class=xl20822214>(Ch&#7919; ký &#273;i&#7879;n t&#7917;,                                                                                      \n");
            htmlStr.Append("   ch&#7919; ký s&#7889;)</td>                                                                                                                                  \n");
            htmlStr.Append("   <td class=xl9122214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append(" <tr height=16 style='mso-height-source:userset;height:15.00pt'>                                                                                                 \n");
            htmlStr.Append("   <td height=16 class=xl6822214 style='height:15.00pt'>&nbsp;</td>                                                                                              \n");
            htmlStr.Append("   <td class=xl6422214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6422214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl7422214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("  </tr>                                                                                                                                                         \n");
            htmlStr.Append("  <tr height=18 style='mso-height-source:userset;height:17.45pt'>                                                                                               \n");
            htmlStr.Append("   <td height=18 class=xl6822214 style='height:17.45pt'>&nbsp;</td>                                                                                             \n");
            htmlStr.Append("   <td class=xl6422214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6422214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6922214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl10422214>&nbsp;</td>                                                                                                                             \n");


            if (dt.Rows[0]["sign_yn"].ToString() == "Y")
            {                                                                                                                                       
                                                                                                                                                                    
            htmlStr.Append(" 		<td colspan=11 class=xl20922214 height=18 width=315 style='border-right:.5pt solid black;                                                               \n");
            htmlStr.Append(" 		  height:17.45pt;width:237pt' align=left valign=top><![if !vml]><span style='mso-ignore:vglayout;                                                       \n");
            htmlStr.Append(" 		  position:absolute;z-index:1;margin-left:154px;margin-top:16px;width:50px;                                                                             \n");
            htmlStr.Append(" 		  height:33px'><img width=50 height=33                                                                                                                  \n");
            htmlStr.Append(" 		  src='D:\\webproject\\e-invoice-ws\\02.Web\\EInvoice\\img\\check_signed.png' v:shapes='Picture_x0020_8'></span><![endif]><span                                \n");
            htmlStr.Append(" 		  style='mso-ignore:vglayout2'>                                                                                                                         \n");
            htmlStr.Append(" 		  <table cellpadding=0 cellspacing=0>                                                                                                                   \n");
            htmlStr.Append(" 		   <tr>                                                                                                                                                 \n");
            htmlStr.Append(" 			<td colspan=11 height=18  width=315 style='height:17.45pt;border-left:none;width:237pt'>Signature                                                   \n");
            htmlStr.Append(" 			Valid</td>                                                                                                                                          \n");
            htmlStr.Append(" 		   </tr>                                                                                                                                                \n");
            htmlStr.Append(" 		  </table>                                                                                                                                              \n");
            htmlStr.Append(" 		  </span></td>                                                                                                                                          \n");
                                                                                                                                                                  
                }else {                                                                                                                                           
                                                                                                                                                                  
            htmlStr.Append(" 		<td colspan=11 class=xl20922214 height=18 width=315 style='border-right:.5pt solid black;                                                               \n");
            htmlStr.Append(" 	height:17.45pt;width:237pt' align=left valign=top></td>                                                                                                     \n");
           	}                                                                                                                                          
                                                                                                                                                             
                                                                                                                                                             
            htmlStr.Append("   <td class=xl7422214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("  </tr>                                                                                                                                                         \n");
            htmlStr.Append("  <tr height=32 style='mso-height-source:userset;height:24.45pt'>                                                                                               \n");
            htmlStr.Append("   <td height=32 class=xl6822214 style='height:24.45pt'>&nbsp;</td>                                                                                             \n");
            htmlStr.Append("   <td class=xl8322214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6422214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6522214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6922214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl10222214 width=25 style='width:23.75pt'>&nbsp;</td>                                                                                                 \n");
            htmlStr.Append("   <td class=xl10322214 width=18 style='width:17.5pt'>&nbsp;</td>                                                                                                 \n");

            if (dt.Rows[0]["sign_yn"].ToString() == "Y")
            {
                htmlStr.Append("   <td colspan=11 class=xl21222214 width=315 style='border-right:.5pt solid black;                                                                              \n");
                htmlStr.Append("   border-left:none;width:237pt'><font class='font1722214'>&#272;&#432;&#7907;c                                                                                 \n");
                htmlStr.Append("   ký b&#7903;i:</font><font class='font1622214'>" + dt.Rows[0]["SignedBy"] + "</font></td>                                                                                  \n");
            }else
            {
                htmlStr.Append("   <td colspan=11 class=xl21222214 width=315 style='border-right:.5pt solid black;                                                                              \n");
                htmlStr.Append("   border-left:none;width:237pt'><font class='font1722214'>&#272;&#432;&#7907;c                                                                                 \n");
                htmlStr.Append("   ký b&#7903;i:</font><font class='font1622214'></font></td>                                                                                  \n");

            }
            htmlStr.Append("   <td class=xl7422214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("  </tr>                                                                                                                                                         \n");
            htmlStr.Append("  <tr height=20 style='mso-height-source:userset;height:18.75pt'>                                                                                                \n");
            htmlStr.Append("   <td height=20 class=xl6822214 style='height:18.75pt'>&nbsp;</td>                                                                                              \n");
            htmlStr.Append("   <td class=xl9622214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6922214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td colspan=15 class=xl18222214>&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("   <td class=xl10522214>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("   <td class=xl12022214 colspan=2>Ngày ký:<span                                                                                                                 \n");
            htmlStr.Append("   style='mso-spacerun:yes'> " + dt.Rows[0]["SignedDate"] + "</span></td>                                                                                                         \n");
            htmlStr.Append("   <td class=xl11822214>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("   <td class=xl11822214>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("   <td class=xl11822214>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("   <td class=xl11822214>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("   <td class=xl11822214>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("   <td class=xl11822214>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("   <td class=xl11822214>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("   <td class=xl11822214>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("   <td class=xl11922214>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("   <td class=xl7422214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("  </tr>                                                                                                                                                         \n");
            htmlStr.Append("  <tr height=22 style='height:21.pt'>                                                                                                                          \n");
            htmlStr.Append("   <td height=22 class=xl6822214 style='height:21.00pt'>&nbsp;</td>                                                                                              \n");
            htmlStr.Append("   <td class=xl6722214 colspan=16><font class='font1122214'>Tra c&#7913;u                                                                                       \n");
            htmlStr.Append("   t&#7841;i Website:</font><font class='font922214'> </font><font                                                                                              \n");
            htmlStr.Append("   class='font1222214'>" + dt.Rows[0]["WEBSITE_EI"] + "</font></td>                                                                                         \n");
            htmlStr.Append("   <td class=xl6922214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl6922214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("   <td class=xl9622214 colspan=3>Mã nh&#7853;n hóa &#273;&#417;n: " + dt.Rows[0]["matracuu"] + "<span                                                                                          \n");
            htmlStr.Append("   style='mso-spacerun:yes'> </span></td>                                                                                                                       \n");
            htmlStr.Append("   <td colspan=8 class=xl18222214></td>                                                                                                                         \n");
            htmlStr.Append("   <td class=xl7422214>&nbsp;</td>                                                                                                                              \n");
            htmlStr.Append("  </tr>                                                                                                                                                         \n");
            htmlStr.Append("  <tr class=xl9522214 height=16 style='height:15.00pt'>                                                                                                          \n");
            htmlStr.Append("   <td colspan=31 height=16 class=xl21522214 width=737 style='border-right:.5pt solid black;                                                                    \n");
            htmlStr.Append("   height:15.00pt;width:556pt'>(C&#7847;n ki&#7875;m tra, &#273;&#7889;i                                                                                         \n");
            htmlStr.Append("   chi&#7871;u khi l&#7853;p, giao, nh&#7853;n hóa &#273;&#417;n)</td>                                                                                          \n");
            htmlStr.Append("  </tr>                                                                                                                                                         \n");
            htmlStr.Append("  <tr class=xl9522214 height=16 style='height:15.00pt'>                                                                                                          \n");
            htmlStr.Append("   <td colspan=31 height=16 class=xl21822214 style='height:15.00pt'>" + dt.Rows[0]["CONTRACT_INFO_EI"] + "</td>                                                                                                                      \n");
            htmlStr.Append("  </tr>                                                                                                                                                         \n");
            htmlStr.Append("  <![if supportMisalignedColumns]>                                                                                                                              \n");
            htmlStr.Append("  <tr height=0 style='display:none'>                                                                                                                            \n");
            htmlStr.Append("   <td width=12 style='width:11.25pt'></td>                                                                                                                         \n");
            htmlStr.Append("   <td width=34 style='width:31.25pt'></td>                                                                                                                        \n");
            htmlStr.Append("   <td width=26 style='width:31.25pt'></td>                                                                                                                        \n");
            htmlStr.Append("   <td width=48 style='width:36pt'></td>                                                                                                                        \n");
            htmlStr.Append("   <td width=4 style='width:3.75pt'></td>                                                                                                                          \n");
            htmlStr.Append("   <td width=41 style='width:31pt'></td>                                                                                                                        \n");
            htmlStr.Append("   <td width=38 style='width:28pt'></td>                                                                                                                        \n");
            htmlStr.Append("   <td width=14 style='width:13.75pt'></td>                                                                                                                        \n");
            htmlStr.Append("   <td width=23 style='width:21.25pt'></td>                                                                                                                        \n");
            htmlStr.Append("   <td width=13 style='width:12.5pt'></td>                                                                                                                        \n");
            htmlStr.Append("   <td width=25 style='width:23.75pt'></td>                                                                                                                        \n");
            htmlStr.Append("   <td width=9 style='width:8.75pt'></td>                                                                                                                          \n");
            htmlStr.Append("   <td width=9 style='width:8.75pt'></td>                                                                                                                          \n");
            htmlStr.Append("   <td width=19 style='width:17.5pt'></td>                                                                                                                        \n");
            htmlStr.Append("   <td width=0></td>                                                                                                                                            \n");
            htmlStr.Append("   <td width=26 style='width:31.25pt'></td>                                                                                                                        \n");
            htmlStr.Append("   <td width=26 style='width:31.25pt'></td>                                                                                                                        \n");
            htmlStr.Append("   <td width=25 style='width:23.75pt'></td>                                                                                                                        \n");
            htmlStr.Append("   <td width=18 style='width:17.5pt'></td>                                                                                                                        \n");
            htmlStr.Append("   <td width=29 style='width:27.5pt'></td>                                                                                                                        \n");
            htmlStr.Append("   <td width=39 style='width:36.25pt'></td>                                                                                                                        \n");
            htmlStr.Append("   <td width=33 style='width:31.25pt'></td>                                                                                                                        \n");
            htmlStr.Append("   <td width=26 style='width:31.25pt'></td>                                                                                                                        \n");
            htmlStr.Append("   <td width=12 style='width:11.25pt'></td>                                                                                                                         \n");
            htmlStr.Append("   <td width=41 style='width:31pt'></td>                                                                                                                        \n");
            htmlStr.Append("   <td width=34 style='width:31.25pt'></td>                                                                                                                        \n");
            htmlStr.Append("   <td width=26 style='width:31.25pt'></td>                                                                                                                        \n");
            htmlStr.Append("   <td width=26 style='width:31.25pt'></td>                                                                                                                        \n");
            htmlStr.Append("   <td width=22 style='width:20.0pt'></td>                                                                                                                        \n");
            htmlStr.Append("   <td width=27 style='width:31.25pt'></td>                                                                                                                        \n");
            htmlStr.Append("   <td width=12 style='width:11.25pt'></td>                                                                                                                         \n");
            htmlStr.Append("  </tr>                                                                                                                                                         \n");
            htmlStr.Append("  <![endif]>                                                                                                                                                    \n");
            htmlStr.Append(" </table>                                                                                                                                                       \n"); 
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