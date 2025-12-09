
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
    public class Solution_New
    {
        public static string View(string tei_einvoice_m_pk, string tei_company_pk, string dbName)
        {
            //dbName = "NOBLANDBD";
            /* string dbUser = "genuwin", dbPwd = "genuwin2";//NOBLANDBD  EINVOICE_252
             string _conString = "Data Source={0};User Id={1};Password={2};Unicode=true";
             _conString = String.Format(_conString, dbName, dbUser, dbPwd);*/
            string _conString = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=123.30.104.243)(PORT=1941))(CONNECT_DATA=(SERVER=dedicated)(SERVICE_NAME=NOBLANDBD)));User ID=genuwin;Password=genuwin2";

            string Procedure = "stacfdstac71_r_02_view_einv_v2";
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

            Procedure = "stacfdstac71_r_03_view_einv";

            command = new OracleCommand(Procedure, connection);
            command.CommandType = CommandType.StoredProcedure;


            command.Parameters.Add("p_tei_einvoice_m_pk", OracleDbType.Varchar2, 1000).Value = tei_einvoice_m_pk;
            command.Parameters.Add("p_user_id", OracleDbType.Varchar2, 1000).Value = "genuwin";
            command.Parameters.Add("p_rtn_value", OracleDbType.RefCursor).Direction = ParameterDirection.Output;
            DataSet ds_d = new DataSet();
            OracleDataAdapter da_d = new OracleDataAdapter(command);
            da_d.Fill(ds_d);
            DataTable dt_d = ds_d.Tables[0];

            int pos = 14, pos_lv = 14, v_count = 0, count_page = 0, count_page_v = 0, r = 0, x = 0;
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
            // if (dt.Rows[0]["CurrencyCodeUSD"].ToString() == "VND")
            // {
            //     read_prive = NumberToTextVN(Decimal.Parse(dt.Rows[0]["TotalAmount"].ToString()));
            // }
            // else
            // {
            //     read_prive = Num2VNText(dt.Rows[0]["TotalAmountInWord"].ToString().ToString(), "USD");
            // }
            // //read_prive = NumberToTextVN(Total_Amount_d);
            // read_prive = read_prive.Replace(",", "");

            // read_prive = read_prive.ToString().Substring(0, 2) + read_prive.ToString().Substring(2, read_prive.Length - 2).ToLower().Replace("mỹ", "Mỹ");

            // read_prive = read_prive.Substring(0, 2) + read_prive.Substring(2, read_prive.Length - 2).ToLower() + ".";
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
            htmlStr.Append("	font-size:13.75pt;                                                                                                                                                            \n");
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
            htmlStr.Append("	font-size:11.5pt;                                                                                                                                                             \n");
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
            htmlStr.Append("	font-size:16.25pt;                                                                                                                                                            \n");
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
            htmlStr.Append("</style>                                                                                                                                                                         \n");
            htmlStr.Append("  </style>                                                                                                                                                                       \n");
            htmlStr.Append("                                                                                                                                                                                 \n");
            htmlStr.Append("                                                                                                                                                                                 \n");
            htmlStr.Append("                                                                                                                                                                                 \n");
            htmlStr.Append("</head>                                                                                                                                                                          \n");
            htmlStr.Append("<body class='A4'>                                                                                                                                                                \n");
            htmlStr.Append("<section class='sheet padding-10mm'>                                                                                                                                             \n");
            htmlStr.Append("	<body link='#0066CC' vlink=purple class=xl70>                                                                                                                                \n");
            htmlStr.Append("                                                                                                                                                                                 \n");
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
            htmlStr.Append("  <td width=32 class=xl66 style='width:30pt;border-top:1.0pt solid windowtext;' align=left valign=top><![if !vml]><span style='mso-ignore:vglayout;                                                                                                            \n");
            htmlStr.Append("  position:absolute;z-index:1;margin-left:1px;margin-top:14px;width:230px;                                                                                                       \n");
            htmlStr.Append("  height:108px'><img width=230 height=87 src='D:/webproject/e-invoice-ws/02.Web/EInvoice/img/Solution.jpg'                                                                             \n");
            htmlStr.Append("  v:shapes='Picture_x0020_1'></span><![endif]><span style='mso-ignore:vglayout2'>	                                                                                             \n");
            htmlStr.Append("	<table cellpadding=0 cellspacing=0>                                                                                                                                          \n");
            htmlStr.Append("   <tr>                                                                                                                                                                          \n");
            htmlStr.Append("    <td height=32  width=32 style='height:30.0pt;width:30pt'>&nbsp;</td>                                                                                               \n");
            htmlStr.Append("   </tr>                                                                                                                                                                         \n");
            htmlStr.Append("  </table>                                                                                                                                                                       \n");
            htmlStr.Append("  </span></td>                                                                                                                                                                   \n");
            htmlStr.Append("  <td class=xl66 width=159 style='width:148.75pt'>&nbsp;</td>                                                                                                                       \n");
            htmlStr.Append("  <td class=xl67 width=80 style='width:75pt'><font class='font6'><b>" + dt.Rows[0]["Seller_Name"] + "  </font><font class='font8'> <i>(" + dt.Rows[0]["Seller_Fname"] + " )</i></b></font></td>                  \n");
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
            htmlStr.Append("  <td colspan=5 class=xl204></font></td>                                                                                                                                         \n");
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
            htmlStr.Append(" <tr height=24 style='mso-height-source:userset;height:22.5pt'>                                                                                                                  \n");
            htmlStr.Append("  <td height=24 class=xl71 style='height:22.5pt'>&nbsp;</td>                                                                                                                     \n");
            htmlStr.Append("  <td class=xl75>Hình thức thanh toán <i>(Method of payment)</i> :</td>                                                                                                          \n");
            htmlStr.Append("  <td colspan=2 class=xl72 style='mso-ignore:colspan'></td>                                                                                                                      \n");
            htmlStr.Append("  <td colspan=2 class=xl199>" + dt.Rows[0]["PaymentMethodCK"] + " </td>                                                                                                                          \n");
            htmlStr.Append("  <td colspan=3 class=xl197>Mã số thuế <i>(Tax code)</i> :</td>                                                                                                                  \n");
            htmlStr.Append("  <td class=xl95></td>                                                                                                                                                           \n");
            htmlStr.Append("  <td class=xl96 width=124 style='width:116.25pt'>" + dt.Rows[0]["BuyerTaxCode"] + " </td>                                                                                                              \n");
            htmlStr.Append("  <td class=xl76>&nbsp;</td>                                                                                                                                                     \n");
            htmlStr.Append(" </tr>                                                                                                                                                                           \n");
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
            htmlStr.Append(" <!--[if gte vml 1]><v:rect id='Rectangle_x0020_1'                                                                                                                               \n");
            htmlStr.Append("   o:spid='_x0000_s1030' style='position:absolute;margin-left:.75pt;                                                                                                             \n");
            htmlStr.Append("   margin-top:.75pt;width:498pt;height:249.75pt;z-index:2;visibility:visible'                                                                                                    \n");
            htmlStr.Append("   strokecolor='#41719c' strokeweight='1pt' o:insetmode='auto'>                                                                                                                  \n");
            htmlStr.Append("   <v:fill src='${pageContext.request.contextPath}/assets/solutionbg1.png' o:title='' opacity='11796f'                                                                           \n");
            htmlStr.Append("    recolor='t' rotate='t' type='frame'/>                                                                                                                                        \n");
            htmlStr.Append("  </v:rect><![endif]--><![if !vml]><span style='mso-ignore:vglayout;position:                                                                                                    \n");
            htmlStr.Append("  absolute;z-index:2;margin-left:0px;margin-top:0px;width:891px;height:357px'><img                                                                                               \n");
            htmlStr.Append("  width=891 height=357 src='D:/webproject/e-invoice-ws/02.Web/EInvoice/img/solutionbg2.png' v:shapes='Rectangle_x0020_1'></span><![endif]>                                            \n");
            htmlStr.Append("  </td>                                                                                                                                                                          \n");
            htmlStr.Append(" </tr>                                                                                                                                                                           \n");
                                                                                                                                                                                       
            
            int l_count = 0;
            if (v_count < 5)
            {
                for (int dtR = 0; dtR < v_count; dtR++)
                {

                    if (dt_d.Rows[dtR][0].ToString().Length > 200)
                    {
                        l_count = dt_d.Rows[dtR][0].ToString().Length / 120;
                    }

                    int stt = dtR + 1;
                    htmlStr.Append(" 			 <tr class=xl115 height=53 style='mso-height-source:userset;height:50.0pt'>                                                                                         \n");
                    htmlStr.Append("			  <td height=53 class=xl106 style='height:50.0pt;border-right:none;border-top:none'>&nbsp;</td>                                                                     \n");
                    htmlStr.Append("			  <td class=xl107 width=32 style='width:30pt' x:num>" + stt + "</td>                                                                                               \n");
                    htmlStr.Append("			  <td colspan=3  class=xl191 width=300 style='border-top:none;border-right:1.0pt solid black;                                                                         \n");
                    htmlStr.Append("			  width:225pt'>&nbsp;" + dt_d.Rows[dtR][0] + "</td>                                                                                                                                    \n");
                    htmlStr.Append("			  <td class=xl108 style='border-top:none;border-left:none'>&nbsp;" + dt_d.Rows[dtR][1] + "</td>                                                                                  \n");
                    htmlStr.Append("			  <td class=xl109 style='border-top:none;border-left:none' >" + dt_d.Rows[dtR][2] + " </td>    \n");
                        if (dt_d.Rows[dtR][3].ToString() == "")
                        {
                            htmlStr.Append("			  <td class=xl110 style='border-top:none;border-top:none' >&nbsp;</td>   \n");
                        }
                        else
                        {
                            htmlStr.Append("			  <td class=xl110 style='border-top:none;border-top:none' >&nbsp;" + dt_d.Rows[dtR][5] + "</td>   \n");
                        }
                                                                                               
                    htmlStr.Append("			  <td class=xl111 style='border-top:none;border-top:none;'>" + dt_d.Rows[dtR][3] + "</td>                                                                                        \n");
                    htmlStr.Append("			  <td class=xl112></td>                                                                                                                                              \n");
                    htmlStr.Append("			  <td class=xl113>" + dt_d.Rows[dtR][4] + "</td>                                                                                                                                 \n");
                    htmlStr.Append("			  <td class=xl114>&nbsp;</td>                                                                                                                                        \n");
                    htmlStr.Append(" 			</tr>                                                                                                                                                                \n");
                }
                if (v_count < 6)
                {
                    for (int i = 0; i < 6 - v_count - l_count; i++)
                    {
                        htmlStr.Append("			 <tr class=xl115 height=53 style='mso-height-source:userset;height:49.9375pt‬'>                                                                                         \n");
                        htmlStr.Append("				  <td height=53 class=xl106 style='height:49.9375pt‬;border-right:none;border-top:none'>&nbsp;</td>                                                                 \n");
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

                    } // for
                }  // if			
            }
            else
            {
                for (int dtR = 0; dtR < v_count; dtR++)
                {
                   
                    if (dt_d.Rows[dtR][0].ToString().Length > 200)
                    {
                        l_count = dt_d.Rows[dtR][0].ToString().Length / 120;
                    }

                    int stt = dtR + 1; //dt_d.Rows[dtR][7]                                                                                                                                                                        
                    htmlStr.Append("	 		 <tr class=xl115 height=53 style='mso-height-source:userset;height:38.6875pt‬'>                                                                                         \n");
                    htmlStr.Append("			  <td height=53 class=xl106 style='height:38.6875pt‬;border-right:none;border-top:none'>&nbsp;</td>                                                                     \n");
                    htmlStr.Append("			  <td class=xl107 width=32 style='width:30pt' x:num>" + stt + "</td>                                                                                               \n");
                    htmlStr.Append("			  <td colspan=3  class=xl191 width=300 style='border-top:none;border-right:1.0pt solid black;                                                                         \n");
                    htmlStr.Append("			  width:225pt'>&nbsp;" + dt_d.Rows[dtR][0] + "</td>                                                                                                                                    \n");
                    htmlStr.Append("			  <td class=xl108 style='border-top:none;border-left:none'>&nbsp;" + dt_d.Rows[dtR][1] + "</td>                                                                                  \n");
                    htmlStr.Append("			  <td class=xl109 style='border-top:none;border-left:none' >" + dt_d.Rows[dtR][2] + " </td>                                                                                      \n");
                    if (dt_d.Rows[dtR][3].ToString() == "")
                    {
                        htmlStr.Append("			  <td class=xl110 style='border-top:none;border-top:none' >&nbsp;</td>   \n");
                    }
                    else
                    {
                        htmlStr.Append("			  <td class=xl110 style='border-top:none;border-top:none' >&nbsp;" + dt_d.Rows[dtR][5] + "</td>   \n");
                    }
                    htmlStr.Append("			  <td class=xl111 style='border-top:none;border-top:none;'>" + dt_d.Rows[dtR][3] + "</td>                                                                                        \n");
                    htmlStr.Append("			  <td class=xl112></td>                                                                                                                                              \n");
                    htmlStr.Append("			  <td class=xl113>" + dt_d.Rows[dtR][4] + "</td>                                                                                                                                 \n");
                    htmlStr.Append("			  <td class=xl114>&nbsp;</td>                                                                                                                                        \n");
                    htmlStr.Append(" 			</tr>                                                                                                                                                                \n");

                }
                if (v_count < 5)
                {
                    for (int i = 0; i < 5 - v_count - l_count; i++)
                    {
                        htmlStr.Append(" <tr class=xl115 height=53 style='mso-height-source:userset;height:38.6875pt‬'>											\n");
                        htmlStr.Append("	  <td height=53 class=xl106 style='height:38.6875pt‬;border-right:none;border-top:none'>&nbsp;</td>                    \n");
                        htmlStr.Append("	  <td class=xl107 width=32 style='width:24pt;border-top:none' x:num></td>                                           \n");
                        htmlStr.Append("	  <td colspan=3  class=xl191 width=300 style='border-top:none;border-right:.5pt solid black;                        \n");
                        htmlStr.Append("	  width:225pt'></td>                                                                                                \n");
                        htmlStr.Append("	  <td class=xl108 style='border-top:none'></td>                                                                     \n");
                        htmlStr.Append("	  <td class=xl109 style='border-top:none' x:num></td>                                                               \n");
                        htmlStr.Append("	  <td class=xl110 style='border-top:none' ></td>                                                                    \n");
                        htmlStr.Append("	  <td class=xl111 style='border-top:none'></td>                                                                     \n");
                        htmlStr.Append("	  <td class=xl112 style='border-top:none'></td>                                                                     \n");
                        htmlStr.Append("	  <td class=xl113 style='border-top:none'></td>                                                                     \n");
                        htmlStr.Append("	  <td class=xl114>&nbsp;</td>                                                                                       \n");
                        htmlStr.Append("	</tr>                                                                                                               \n");

                    } // for
                }  // if	
            }

            htmlStr.Append(" <tr class=xl144 height=21 style='mso-height-source:userset;height:19.6875pt'>                                                                                                     \n");
            htmlStr.Append("  <td height=21 class=xl138 style='height:19.6875pt'>&nbsp;</td>                                                                                                                   \n");
            htmlStr.Append("  <td class=xl139>*</td>                                                                                                                                                         \n");
            htmlStr.Append("  <td colspan=3 class=xl188 width=300 style='border-right:1.0pt solid black;                                                                                                      \n");
            htmlStr.Append("  width:225pt'>" + dt.Rows[0]["exchangerate_no"] + "</td>                                                                                                                                             \n");
            htmlStr.Append("  <td class=xl140 width=54 style='border-left:none;width:51.25pt'>&nbsp;</td>                                                                                                       \n");
            htmlStr.Append("  <td class=xl140 width=61 style='border-left:none;width:57.5pt'>&nbsp;</td>                                                                                                       \n");
            htmlStr.Append("  <td class=xl141 width=33 style='width:31.25pt'>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append("  <td class=xl141 width=82 style='width:77.5pt'>&nbsp;</td>                                                                                                                        \n");
            htmlStr.Append("  <td class=xl141 width=12 style='width:11.25pt'>&nbsp;</td>                                                                                                                         \n");
            htmlStr.Append("  <td class=xl142 width=124 style='width:116.25pt'>&nbsp;</td>                                                                                                                       \n");
            htmlStr.Append("  <td class=xl143>&nbsp;</td>                                                                                                                                                    \n");
            htmlStr.Append(" </tr>                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                                 \n");
            htmlStr.Append(" <tr height=20 style='mso-height-source:userset;height:18.75pt'>                                                                                                                  \n");
            htmlStr.Append("  <td height=20 class=xl149 style='height:18.75pt'>&nbsp;</td>                                                                                                                    \n");
            htmlStr.Append("  <td class=xl150></td>                                                                                                                                                          \n");
            htmlStr.Append("  <td colspan=3 class=xl151 style='mso-ignore:colspan'></td>                                                                                                                     \n");
            htmlStr.Append("  <td class=xl75 colspan=4 style='mso-ignore:colspan'>Cộng tiền hàng <i>(Net total)</i> :</font></td>                                                                            \n");
            htmlStr.Append("  <td class=xl152></td>                                                                                                                                                          \n");
            htmlStr.Append("  <td class=xl153>" + dt.Rows[0]["netamount_display"] + " </td>                                                                                                                                           \n");
            htmlStr.Append("  <td class=xl76>&nbsp;</td>                                                                                                                                                     \n");
            htmlStr.Append(" </tr>                                                                                                                                                                           \n");
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
           
            if(dt.Rows[0]["vatamount_display"].ToString().Trim() == "0")
            {
                htmlStr.Append("  <td class=xl155>-</td>                                                                                                                                           \n");

            }
            else
            {
                htmlStr.Append("  <td class=xl155>" + dt.Rows[0]["vatamount_display"] + " </td>                                                                                                                                           \n");

            }




            htmlStr.Append("  <td class=xl76>&nbsp;</td>                                                                                                                                                     \n");
            htmlStr.Append(" </tr>                                                                                                                                                                           \n");
            htmlStr.Append(" <tr height=20 style='mso-height-source:userset;height:18.75pt'>                                                                                                                  \n");
            htmlStr.Append("  <td height=20 class=xl149 style='height:18.75pt'>&nbsp;</td>                                                                                                                    \n");
            htmlStr.Append("  <td class=xl150></td>                                                                                                                                                          \n");
            htmlStr.Append("  <td colspan=3 class=xl151 style='mso-ignore:colspan'></td>                                                                                                                     \n");
            htmlStr.Append("  <td class=xl75 colspan=4 style='mso-ignore:colspan'>T&#7893;ng ti&#7873;n                                                                                                      \n");
            htmlStr.Append("  thanh toán (<font class='font45'>Gross total</font><font class='font46'>) :</font></td>                                                                                        \n");
            htmlStr.Append("  <td class=xl152></td>                                                                                                                                                          \n");
            htmlStr.Append("  <td class=xl155>" + dt.Rows[0]["totalamount_display"] + " </td>                                                                                                                                    \n");
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
            htmlStr.Append(" <tr class=xl167 height=31 style='mso-height-source:userset;height:29.0625pt'>                                                                                                     \n");
            htmlStr.Append("  <td height=31 class=xl165 style='height:29.0625pt;border-top:1.0pt solid black;'>&nbsp;</td>                                                                                      \n");
            htmlStr.Append("  <td colspan=2 class=xl186 width=191 style='width:143pt;border-top:1.0pt solid black;'>Người mua (Buyer)</td>                                                                    \n");
            htmlStr.Append("  <td colspan=2 class=xl186 width=141 style='width:106pt;border-top:1.0pt solid black;'></td>                                                                                     \n");
            htmlStr.Append("  <td colspan=3 class=xl186 width=148 style='width:112pt;border-top:1.0pt solid black;'></td>                                                                                     \n");
            htmlStr.Append("  <td colspan=3 class=xl187 width=218 style='width:164pt;border-top:1.0pt solid black;'>Ng&#432;&#7901;i bán                                                                      \n");
            htmlStr.Append("  (Seller)</td>                                                                                                                                                                  \n");
            htmlStr.Append("  <td class=xl166 style='border-top:1.0pt solid black;'>&nbsp;</td>                                                                                                               \n");
            htmlStr.Append(" </tr>                                                                                                                                                                           \n");
            htmlStr.Append(" <tr height=17 style='height:12.75pt'>                                                                                                                                           \n");
            htmlStr.Append("  <td height=17 class=xl71 style='height:12.75pt'>&nbsp;</td>                                                                                                                    \n");
            htmlStr.Append("  <td colspan=2 class=xl177 width=191 style='width:143pt'>Ký, ghi rõ h&#7885;                                                                                                    \n");
            htmlStr.Append("  tên</td>                                                                                                                                                                       \n");
            htmlStr.Append("  <td colspan=2 class=xl177 width=141 style='width:106pt'></td>                                                                                                                  \n");
            htmlStr.Append("  <td colspan=3 class=xl177 width=148 style='width:112pt'></td>                                                                                                                  \n");
            htmlStr.Append("  <td colspan=3 class=xl177 width=218 style='width:164pt'>Ký, &#273;óng                                                                                                         \n");
            htmlStr.Append("  d&#7845;u, ghi rõ h&#7885; tên</td>                                                                                                                                            \n");
            htmlStr.Append("  <td class=xl76>&nbsp;</td>                                                                                                                                                     \n");
            htmlStr.Append(" </tr>                                                                                                                                                                           \n");
            htmlStr.Append(" <tr height=17 style='height:12.75pt'>                                                                                                                                           \n");
            htmlStr.Append("  <td height=17 class=xl71 style='height:12.75pt'>&nbsp;</td>                                                                                                                    \n");
            htmlStr.Append("  <td colspan=2 class=xl177 width=191 style='width:143pt'>(Sign &amp; full                                                                                                       \n");
            htmlStr.Append("  name)</td>                                                                                                                                                                     \n");
            htmlStr.Append("  <td colspan=2 class=xl177 width=141 style='width:106pt'></td>                                                                                                                  \n");
            htmlStr.Append("  <td colspan=3 class=xl177 width=148 style='width:112pt'></td>                                                                                                                  \n");
            htmlStr.Append("  <td colspan=3 class=xl177 width=218 style='width:164pt'>(Sign, stamp &amp;                                                                                                     \n");
            htmlStr.Append("  full name)</td>                                                                                                                                                                \n");
            htmlStr.Append("  <td class=xl76>&nbsp;</td>                                                                                                                                                     \n");
            htmlStr.Append(" </tr>                                                                                                                                                                           \n");
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
            htmlStr.Append("    <td height=20 class=xl71 style='height:18.75pt'>&nbsp;</td>                                                                                                                                                                             \n");
            htmlStr.Append("    <td class=xl169></td>                                                                                                                                                                                                                  \n");
            htmlStr.Append("    <td colspan=4 class=xl170 style='mso-ignore:colspan'></td>                                                                                                                                                                             \n");
            htmlStr.Append("    <td colspan=5 class=xl178 width=312 style='border-right:.5pt solid #0066CC;                                                                                                                                                            \n");
            htmlStr.Append("    width:235pt' x:str='Signature Valid'>Signature Valid<span                                                                                                                                                                             \n");
            htmlStr.Append("    style='mso-spacerun:yes'> </span></td>                                                                                                                                                                                                 \n");
            htmlStr.Append("    <td class=xl76>&nbsp;</td>                                                                                                                                                                                                             \n");
            htmlStr.Append("   </tr>                                                                                                                                                                                                                                   \n");
            htmlStr.Append("   <tr height=36 style='mso-height-source:userset;height:27.0pt'>                                                                                                                                                                          \n");
            htmlStr.Append("    <td height=36 class=xl71 style='height:27.0pt'>&nbsp;</td>                                                                                                                                                                             \n");
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
            htmlStr.Append("    href=" + dt.Rows[0]["tracuuwebsite"] + "><span style='color:#003366;                                                                                                                                                                                       \n");
            htmlStr.Append("    font-style:italic'>" + dt.Rows[0]["tracuuwebsite"] + "</span></a></td>                                                                                                                                                                                     \n");
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

    }
}