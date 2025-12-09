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
    public class Yupoong_2_New_C
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

            //read_prive = dt.Rows[0]["amount_word_vie"].ToString();

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
            read_prive = dt.Rows[0]["AMOUNT_WORD_VIE"].ToString();//
            int end = 0;
            int count = count_page_v + r;
            double height = 130;
            StringBuilder htmlStr = new StringBuilder("");
            string heigh = "", heigh_d = "";

            /////////////////////////////
            htmlStr.Append("    <!DOCTYPE html PUBLIC '-//W3C//DTD HTML 4.01 Transitional//EN' 'http://www.w3.org/TR/html4/loose.dtd'>                                                               \n");
            htmlStr.Append("    <html>                                                               \n");
            htmlStr.Append("    <head>                                                               \n");
            htmlStr.Append("    <meta http-equiv='Content-Type' content='text/html;charset=UTF-8'>                                                               \n");
            htmlStr.Append("                                                                   \n");
            htmlStr.Append("    <script type='text/javascript' src='${pageContext.request.contextPath}/system/syscommand.js'></script>                                                               \n");
            htmlStr.Append("    <title>Report E-Invoice</title>                                                               \n");
            htmlStr.Append("     <!-- Normalize or reset CSS with your favorite library -->                                                               \n");
            htmlStr.Append("      <link rel='stylesheet' href='https://cdnjs.cloudflare.com/ajax/libs/normalize/3.0.3/normalize.css'>                                                               \n");
            htmlStr.Append("                                                                   \n");
            htmlStr.Append("      <!-- Load paper.css for happy printing -->                                                               \n");
            htmlStr.Append("      <link rel='stylesheet' href='https://cdnjs.cloudflare.com/ajax/libs/paper-css/0.2.3/paper.css'>                                                               \n");
            htmlStr.Append("                                                                   \n");
            htmlStr.Append("      <!-- Set page size here: A5, A4 or A3 -->                                                               \n");
            htmlStr.Append("      <!-- Set also 'landscape' if you need -->                                                               \n");
            htmlStr.Append("      <style>@page { size: A4 }</style>                                                               \n");
            htmlStr.Append("      <link href='https://fonts.googleapis.com/css?family=Tangerine:700' rel='stylesheet' type='text/css'>                                                               \n");
            htmlStr.Append("      <style>                                                               \n");
            htmlStr.Append("        /*body   { font-family: serif }                                                               \n");
            htmlStr.Append("        h1     { font-family: 'Tangerine', cursive; font-size: 40pt; line-height: 18mm}                                                               \n");
            htmlStr.Append("        h2, h3 { font-family: 'Tangerine', cursive; font-size: 24pt; line-height: 7mm }                                                               \n");
            htmlStr.Append("        h4     { font-size: 13pt; line-height: 1mm }                                                               \n");
            htmlStr.Append("        h2 + p { font-size: 18pt; line-height: 7mm }                                                               \n");
            htmlStr.Append("        h3 + p { font-size: 14pt; line-height: 7mm }                                                               \n");
            htmlStr.Append("        li     { font-size: 11pt; line-height: 5mm }                                                               \n");
            htmlStr.Append("                                                                   \n");
            htmlStr.Append("        h1      { margin: 0 }                                                               \n");
            htmlStr.Append("        h1 + ul { margin: 2mm 0 5mm }                                                               \n");
            htmlStr.Append("        h2, h3  { margin: 0 3mm 3mm 0; float: left }                                                               \n");
            htmlStr.Append("        h2 + p,                                                               \n");
            htmlStr.Append("        h3 + p  { margin: 0 0 3mm 50mm }                                                               \n");
            htmlStr.Append("        //h4      { margin: 1mm 0 0 2mm; border-bottom: 1px solid black }                                                               \n");
            htmlStr.Append("        h4 + ul { margin: 5mm 0 0 50mm }                                                               \n");
            htmlStr.Append("        article { border: 4px double black; padding: 5mm 10mm; border-radius: 3mm }*/                                                               \n");
            htmlStr.Append("        body {                                                               \n");
            htmlStr.Append("           		 color: blue;                                                               \n");
            htmlStr.Append("           		 font-size:100%;                                                               \n");
            htmlStr.Append("           		 background-image: url('assets/Solution.jpg');                                                               \n");
            htmlStr.Append("    		 }                                                               \n");
            htmlStr.Append("    	h1 {                                                               \n");
            htmlStr.Append("    	        color: #00FF00;                                                               \n");
            htmlStr.Append("    	}                                                               \n");
            htmlStr.Append("    	p {                                                               \n");
            htmlStr.Append("    	        color: rgb(0,0,255)                                                               \n");
            htmlStr.Append("    	}                                                               \n");
            htmlStr.Append("    	                                                               \n");
            htmlStr.Append("       headline1 {                                                               \n");
            htmlStr.Append("          background-image: url(assets/Solution.jpg);                                                               \n");
            htmlStr.Append("          background-repeat: no-repeat;                                                               \n");
            htmlStr.Append("          background-position: left top;                                                               \n");
            htmlStr.Append("          padding-top:68px;                                                               \n");
            htmlStr.Append("          margin-bottom:50px;                                                               \n");
            htmlStr.Append("       }                                                               \n");
            htmlStr.Append("       headline2 {                                                               \n");
            htmlStr.Append("          background-image: url(images/newsletter_headline2.gif);                                                               \n");
            htmlStr.Append("          background-repeat: no-repeat;                                                               \n");
            htmlStr.Append("          background-position: left top;                                                               \n");
            htmlStr.Append("          padding-top:68px;                                                               \n");
            htmlStr.Append("       }                                                               \n");
            htmlStr.Append("    <!--table                                                               \n");
            htmlStr.Append("    	{mso-displayed-decimal-separator:'\\.';                                                               \n");
            htmlStr.Append("    	mso-displayed-thousand-separator:'\\, ';}                                                               \n");

            htmlStr.Append("    .font520255                                                               \n");
            htmlStr.Append("    	{color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:12.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font620255                                                               \n");
            htmlStr.Append("    	{color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:10.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font720255                                                               \n");
            htmlStr.Append("    	{color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:12.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font820255                                                               \n");
            htmlStr.Append("    	{color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:11.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font920255                                                               \n");
            htmlStr.Append("    	{color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:11.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font1020255                                                               \n");
            htmlStr.Append("    	{color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:10.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font1120255                                                               \n");
            htmlStr.Append("    	{color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:11.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:italic;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font1220255                                                               \n");
            htmlStr.Append("    	{color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:10.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:italic;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font1320255                                                               \n");
            htmlStr.Append("    	{color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:11.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:italic;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font1420255                                                               \n");
            htmlStr.Append("    	{color:#993300;                                                               \n");
            htmlStr.Append("    	font-size:11.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font1520255                                                               \n");
            htmlStr.Append("    	{color:#333399;                                                               \n");
            htmlStr.Append("    	font-size:12.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font1620255                                                               \n");
            htmlStr.Append("    	{color:red;                                                               \n");
            htmlStr.Append("    	font-size:11.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font1720255                                                               \n");
            htmlStr.Append("    	{color:red;                                                               \n");
            htmlStr.Append("    	font-size:11.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:italic;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font1820255                                                               \n");
            htmlStr.Append("    	{color:red;                                                               \n");
            htmlStr.Append("    	font-size:11.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font1920255                                                               \n");
            htmlStr.Append("    	{color:red;                                                               \n");
            htmlStr.Append("    	font-size:10.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font2020255                                                               \n");
            htmlStr.Append("    	{color:red;                                                               \n");
            htmlStr.Append("    	font-size:10.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:italic;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font2120255                                                               \n");
            htmlStr.Append("    	{color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:11.5pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font2220255                                                               \n");
            htmlStr.Append("    	{color:#0066CC;                                                               \n");
            htmlStr.Append("    	font-size:10.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font2320255                                                               \n");
            htmlStr.Append("    	{color:red;                                                               \n");
            htmlStr.Append("    	font-size:17.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font2420255                                                               \n");
            htmlStr.Append("    	{color:red;                                                               \n");
            htmlStr.Append("    	font-size:17.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:italic;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font517145                                                               \n");
            htmlStr.Append("    	{color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:36.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font617145                                                               \n");
            htmlStr.Append("    	{color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:20.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font717145                                                               \n");
            htmlStr.Append("    	{color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:19.5pt;                                                             \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font817145                                                               \n");
            htmlStr.Append("    	{color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:19.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font917145                                                               \n");
            htmlStr.Append("    	{color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font1017145                                                               \n");
            htmlStr.Append("    	{color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:20.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font1117145                                                               \n");
            htmlStr.Append("    	{color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:italic;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font1217145                                                               \n");
            htmlStr.Append("    	{color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:italic;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font1317145                                                               \n");
            htmlStr.Append("    	{color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:italic;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font1417145                                                               \n");
            htmlStr.Append("    	{color:#993300;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font1517145                                                               \n");
            htmlStr.Append("    	{color:#333399;                                                               \n");
            htmlStr.Append("    	font-size:22.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font1617145                                                               \n");
            htmlStr.Append("    	{color:red;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font1717145                                                               \n");
            htmlStr.Append("    	{color:red;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:italic;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font1817145                                                               \n");
            htmlStr.Append("    	{color:red;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font1917145                                                               \n");
            htmlStr.Append("    	{color:red;                                                               \n");
            htmlStr.Append("    	font-size:14.5pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	white-space:normal;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font2017145                                                               \n");
            htmlStr.Append("    	{color:red;                                                               \n");
            htmlStr.Append("    	font-size:20.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:italic;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("        white - space:normal;   \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font2117145                                                               \n");
            htmlStr.Append("    	{color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:18.0pt;                                                             \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font2217145                                                               \n");
            htmlStr.Append("    	{color:#0066CC;                                                               \n");
            htmlStr.Append("    	font-size:20.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font2317145                                                               \n");
            htmlStr.Append("    	{color:red;                                                               \n");
            htmlStr.Append("    	font-size:25.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .font2417145                                                               \n");
            htmlStr.Append("    	{color:red;                                                               \n");
            htmlStr.Append("    	font-size:25.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:italic;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;}                                                               \n");
            htmlStr.Append("    .xl6517145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:20.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	mso-background-source:auto;                                                               \n");
            htmlStr.Append("    	mso-pattern:auto;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl6617145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:20.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl6717145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:20.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl6817145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:20.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl6917145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:20.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl7017145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:20.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl7117145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl7217145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:20.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl7317145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size: 20pt;                                                             \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl7417145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:20.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:normal;}                                                               \n");
            htmlStr.Append("    .xl7517145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:19.5pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl7617145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:20.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:'\\@';                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl7717145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:20.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl7817145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:20.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl7917145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:25.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:right;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl8017145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:20.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-left:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl8117145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size: 20pt;                                                             \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl8217145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size: 20pt;                                                             \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl8317145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:20.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:italic;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl8417145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl8517145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl8617145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:.5pt dotted windowtext;                                                                    \n");
            htmlStr.Append("    	border-left:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;                                                               \n");
            htmlStr.Append("    	mso-text-control:shrinktofit;}                                                               \n");
            htmlStr.Append("    .xl8717145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size: 20pt;                                                             \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:right;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:.5pt dotted windowtext;                                                                   \n");
            htmlStr.Append("    	border-left:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;                                                               \n");
            htmlStr.Append("    	mso-text-control:shrinktofit;}                                                               \n");
            htmlStr.Append("    .xl8817145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:'_\\(* \\#\\,\\#\\#0_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;                                                               \n");
            htmlStr.Append("    	mso-text-control:shrinktofit;}                                                               \n");
            htmlStr.Append("    .xl8917145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size: 20pt;                                                             \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:'_\\(* \\#\\,\\#\\#0_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                                \n");
            htmlStr.Append("    	text-align:right;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:.5pt dotted windowtext;                                                                    \n");
            htmlStr.Append("    	border-left:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;                                                               \n");
            htmlStr.Append("    	mso-text-control:shrinktofit;}                                                               \n");
            htmlStr.Append("    .xl9017145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:'_\\(* \\#\\,\\#\\#0_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                              \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;                                                               \n");
            htmlStr.Append("    	mso-text-control:shrinktofit;}                                                               \n");
            htmlStr.Append("    .xl9117145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:'_\\(* \\#\\,\\#\\#0_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                              \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:.5pt dotted windowtext;                                                                   \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;                                                               \n");
            htmlStr.Append("    	mso-text-control:shrinktofit;}                                                               \n");
            htmlStr.Append("    .xl9217145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size: 20pt;                                                             \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:right;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:.5pt dotted windowtext;                                                                \n");
            htmlStr.Append("    	border-left:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("        font-style: normal;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl9317145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size: 20pt;                                                             \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:italic;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl9417145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size: 20pt;                                                             \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt dotted windowtext;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:.5pt dotted windowtext;                                                               \n");
            htmlStr.Append("    	border-left:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;                                                               \n");
            htmlStr.Append("    	mso-text-control:shrinktofit;}                                                               \n");
            htmlStr.Append("    .xl9517145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size: 20pt;                                                             \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:'\\@';                                                               \n");
            htmlStr.Append("    	text-align:right;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt dotted windowtext;                                                              \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:.5pt dotted windowtext;                                                              \n");
            htmlStr.Append("    	border-left:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;                                                               \n");
            htmlStr.Append("    	mso-text-control:shrinktofit;}                                                               \n");
            htmlStr.Append("    .xl9617145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size: 20pt;                                                             \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:'\\@';                                                               \n");
            htmlStr.Append("    	text-align:right;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt dotted windowtext;                                                                \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:.5pt dotted windowtext; ;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;                                                               \n");
            htmlStr.Append("    	mso-text-control:shrinktofit;}                                                               \n");
            htmlStr.Append("    .xl9717145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:'\\@';                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;                                                               \n");
            htmlStr.Append("    	mso-text-control:shrinktofit;}                                                               \n");
            htmlStr.Append("    .xl9817145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:'\\@';                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;                                                               \n");
            htmlStr.Append("    	mso-text-control:shrinktofit;}                                                               \n");
            htmlStr.Append("    .xl9917145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:'\\@';                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;                                                               \n");
            htmlStr.Append("    	mso-text-control:shrinktofit;}                                                               \n");
            htmlStr.Append("    .xl10017145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:'\\@';                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;                                                               \n");
            htmlStr.Append("    	mso-text-control:shrinktofit;}                                                               \n");
            htmlStr.Append("    .xl10117145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:'\\@';                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt dotted windowtext;                                                             \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:.5pt dotted windowtext;                                                               \n");
            htmlStr.Append("    	border-left:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:normal;}                                                               \n");
            htmlStr.Append("    .xl10217145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:'\\@';                                                               \n");
            htmlStr.Append("    	text-align:right;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt dotted windowtext;                                                              \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                            \n");
            htmlStr.Append("    	border-bottom:.5pt dotted windowtext;                                                              \n");
            htmlStr.Append("    	border-left:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:normal;}                                                               \n");
            htmlStr.Append("    .xl10317145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:'\\@';                                                               \n");
            htmlStr.Append("    	text-align:right;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:normal;}                                                               \n");
            htmlStr.Append("    .xl10417145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:'\\@';                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt dotted windowtext;                                                              \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:.5pt dotted windowtext;                                                              \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:normal;}                                                               \n");
            htmlStr.Append("    .xl10517145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:'\\@';                                                               \n");
            htmlStr.Append("    	text-align:left;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt dotted windowtext;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:.5pt dotted windowtext;                                                              \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:normal;}                                                               \n");
            htmlStr.Append("    .xl10617145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size: 20pt;                                                             \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl10717145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size: 20pt;                                                             \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:'\\@';                                                               \n");
            htmlStr.Append("    	text-align:right;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;                                                               \n");
            htmlStr.Append("    	mso-text-control:shrinktofit;}                                                               \n");
            htmlStr.Append("    .xl10817145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:normal;}                                                               \n");
            htmlStr.Append("    .xl10917145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl11017145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:normal;}                                                               \n");
            htmlStr.Append("    .xl11117145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:normal;}                                                               \n");
            htmlStr.Append("    .xl11217145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size: 20pt;                                                             \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:normal;}                                                               \n");
            htmlStr.Append("    .xl11317145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:20.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl11417145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl11517145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:normal;}                                                               \n");
            htmlStr.Append("    .xl11617145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:italic;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:normal;}                                                               \n");
            htmlStr.Append("    .xl11717145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:20.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:italic;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl11817145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size: 20pt;                                                             \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:italic;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:normal;}                                                               \n");
            htmlStr.Append("    .xl11917145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:20.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:italic;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl12017145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size: 20pt;                                                             \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:normal;}                                                               \n");
            htmlStr.Append("    .xl12117145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:#C00000;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl12217145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size: 20pt;                                                             \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:normal;}                                                               \n");
            htmlStr.Append("    .xl12317145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt dotted windowtext;                                                                  \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:.5pt dotted windowtext;                                                                  \n");
            htmlStr.Append("    	border-left:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:normal;}                                                               \n");
            htmlStr.Append("    .xl12417145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:normal;}                                                               \n");
            htmlStr.Append("    .xl12517145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size: 20pt;                                                             \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	border-top:.5pt dotted windowtext;                                                                  \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:.5pt dotted windowtext;                                                                  \n");
            htmlStr.Append("    	text-align:left;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:normal;}                                                               \n");
            htmlStr.Append("    .xl12617145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:'\\@';                                                               \n");
            htmlStr.Append("    	text-align:right;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;                                                               \n");
            htmlStr.Append("    	mso-text-control:shrinktofit;}                                                               \n");
            htmlStr.Append("    .xl12717145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-left:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl12817145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:normal;}                                                               \n");
            htmlStr.Append("    .xl12917145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:36.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl13017145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:36.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl13117145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size: 20pt;                                                             \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:left;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt dotted windowtext;                                                                   \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:.5pt dotted windowtext;                                                                   \n");
            htmlStr.Append("    	border-left:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:normal;}                                                               \n");
            htmlStr.Append("    .xl13217145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size: 20pt;                                                             \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:left;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:normal;}                                                               \n");
            htmlStr.Append("    .xl13317145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:red;                                                               \n");
            htmlStr.Append("    	font-size:20.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:italic;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:bottom;                                                               \n");
            htmlStr.Append("    	border-top:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl13417145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:25.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:bottom;                                                               \n");
            htmlStr.Append("    	border-top:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl13517145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:25.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:bottom;                                                               \n");
            htmlStr.Append("    	border-top:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl13617145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:25.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:bottom;                                                               \n");
            htmlStr.Append("    	border-top:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl13717145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:red;                                                               \n");
            htmlStr.Append("    	font-size:20.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:italic;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl13817145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:25.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:bottom;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl13917145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:red;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:normal;}                                                               \n");
            htmlStr.Append("    .xl14017145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:red;                                                               \n");
            htmlStr.Append("    	font-size:20.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:italic;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-left:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl14117145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:red;                                                               \n");
            htmlStr.Append("    	font-size:20.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:italic;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl14217145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:red;                                                               \n");
            htmlStr.Append("    	font-size:20.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:italic;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl14317145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:#0070C0;                                                               \n");
            htmlStr.Append("    	font-size:20.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl14417145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:36.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl14517145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:left;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;                                                               \n");
            htmlStr.Append("    	mso-text-control:shrinktofit;}                                                               \n");
            htmlStr.Append("    .xl14617145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:#002060;                                                               \n");
            htmlStr.Append("    	font-size:36.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl14717145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:#002060;                                                               \n");
            htmlStr.Append("    	font-size:36.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl14817145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:red;                                                               \n");
            htmlStr.Append("    	font-size:26.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl14917145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size: 20pt;                                                             \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:normal;}                                                               \n");
            htmlStr.Append("    .xl15017145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size: 20pt;                                                             \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl15117145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size: 20pt;                                                             \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl15217145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:20.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:italic;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-left:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl15317145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:19.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:italic;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl15417145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:20.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:italic;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl15517145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:20.0pt;                                                              \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl15617145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size: 20pt;                                                             \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl15717145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:20.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:left;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl15817145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-left:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl15917145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl16017145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	font-size: 20pt;                                                             \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:left;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:.5pt dotted windowtext;                                                                   \n");
            htmlStr.Append("    	border-left:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:normal;}                                                               \n");
            htmlStr.Append("    .xl16117145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size: 20pt;                                                             \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:left;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:normal;}                                                               \n");
            htmlStr.Append("    .xl16217145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size: 20pt;                                                             \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:left;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:normal;}                                                               \n");
            htmlStr.Append("    .xl16317145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:20.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:italic;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl16417145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:20.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:italic;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl16517145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size: 20pt;                                                             \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:italic;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl16617145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:20.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl16717145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:'\\@';                                                               \n");
            htmlStr.Append("    	text-align:right;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;                                                               \n");
            htmlStr.Append("    	mso-text-control:shrinktofit;}                                                               \n");
            htmlStr.Append("    .xl16817145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:'\\@';                                                               \n");
            htmlStr.Append("    	text-align:right;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;                                                               \n");
            htmlStr.Append("    	mso-text-control:shrinktofit;}                                                               \n");
            htmlStr.Append("    .xl16917145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:left;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl17017145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:red;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:left;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;                                                               \n");
            htmlStr.Append("    	mso-text-control:shrinktofit;}                                                               \n");
            htmlStr.Append("    .xl17117145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:red;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:left;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;                                                               \n");
            htmlStr.Append("    	mso-text-control:shrinktofit;}                                                               \n");
            htmlStr.Append("    .xl17217145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:red;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:left;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;                                                               \n");
            htmlStr.Append("    	mso-text-control:shrinktofit;}                                                               \n");
            htmlStr.Append("    .xl17317145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size: 20pt;                                                             \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:italic;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:normal;}                                                               \n");
            htmlStr.Append("    .xl17417145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size: 20pt;                                                             \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:normal;}                                                               \n");
            htmlStr.Append("    .xl17517145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size: 20pt;                                                             \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:left;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:normal;}                                                               \n");
            htmlStr.Append("    .xl17617145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:left;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;                                                               \n");
            htmlStr.Append("    	mso-text-control:shrinktofit;}                                                               \n");
            htmlStr.Append("    .xl17717145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:left;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;                                                               \n");
            htmlStr.Append("    	mso-text-control:shrinktofit;}                                                               \n");
            htmlStr.Append("    .xl17817145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:21.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:italic;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl17917145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size: 20pt;                                                             \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl18017145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:25.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    .xl18117145                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:20.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:700;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    	.xl70171451                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:1.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:none;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    	.xl166171451                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:1.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:center;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:nowrap;}                                                               \n");
            htmlStr.Append("    	.xl120171451                                                               \n");
            htmlStr.Append("    	{padding:0px;                                                               \n");
            htmlStr.Append("    	mso-ignore:padding;                                                               \n");
            htmlStr.Append("    	color:windowtext;                                                               \n");
            htmlStr.Append("    	font-size:1.0pt;                                                               \n");
            htmlStr.Append("    	font-weight:400;                                                               \n");
            htmlStr.Append("    	font-style:normal;                                                               \n");
            htmlStr.Append("    	text-decoration:none;                                                               \n");
            htmlStr.Append("    	font-family:'Times New Roman', serif;                                                               \n");
            htmlStr.Append("    	mso-font-charset:0;                                                               \n");
            htmlStr.Append("    	mso-number-format:General;                                                               \n");
            htmlStr.Append("    	text-align:general;                                                               \n");
            htmlStr.Append("    	vertical-align:middle;                                                               \n");
            htmlStr.Append("    	border-top:none;                                                               \n");
            htmlStr.Append("    	border-right:.5pt solid windowtext;                                                               \n");
            htmlStr.Append("    	border-bottom:none;                                                               \n");
            htmlStr.Append("    	border-left:none;                                                               \n");
            htmlStr.Append("    	background:white;                                                               \n");
            htmlStr.Append("    	mso-pattern:black none;                                                               \n");
            htmlStr.Append("    	white-space:normal;}                                                               \n");
            htmlStr.Append("    -->                                                               \n");
            htmlStr.Append("    </style>                                                               \n");
            htmlStr.Append("    </head>                                                               \n");
            htmlStr.Append("    <body class='A4'>                                                               \n");
            //////////////////////////////
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

            double v_totalHeightLastPage = 500;// 258.5;

            double v_totalHeightPage = 580;//   540;

            for (int k = 0; k < v_countNumberOfPages; k++)
            {
                v_totalHeightPage = 750;// 540;

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
                /////////////////////////////
                //html part 2 (from top to 1st line)
                htmlStr.Append("    <table border=0 cellpadding=0 cellspacing=0 width=642 class=xl6617145                                                               \n");
                htmlStr.Append("    					 style='border-collapse:collapse;table-layout:fixed;width:1058pt'>                                                               \n");
                htmlStr.Append("    					 <col class=xl6617145 width=6 style='mso-width-source:userset;mso-width-alt:                                                               \n");
                htmlStr.Append("    					 199;width:7pt'>                                                               \n");
                htmlStr.Append("    					 <col class=xl6617145 width=33 style='mso-width-source:userset;mso-width-alt:                                                               \n");
                htmlStr.Append("    					 1166;width:42pt'>                                                               \n");
                htmlStr.Append("    					 <col class=xl6617145 width=70 style='mso-width-source:userset;mso-width-alt:                                                               \n");
                htmlStr.Append("    					 2474;width:125pt'>                                                               \n");
                htmlStr.Append("    					 <col class=xl6617145 width=55 style='mso-width-source:userset;mso-width-alt:                                                               \n");
                htmlStr.Append("    					 1962;width:160pt'>                                                               \n");
                htmlStr.Append("    					 <col class=xl6617145 width=41 style='mso-width-source:userset;mso-width-alt:                                                               \n");
                htmlStr.Append("    					 1450;width:140pt'>                                                               \n");
                htmlStr.Append("    					 <col class=xl6617145 width=98 style='mso-width-source:userset;mso-width-alt:                                                               \n");
                htmlStr.Append("    					 3498;width:61pt'>                                                               \n");
                htmlStr.Append("    					 <col class=xl6617145 width=27 style='mso-width-source:userset;mso-width-alt:                                                               \n");
                htmlStr.Append("    					 967;width:13pt'>                                                               \n");
                htmlStr.Append("    					 <col class=xl6617145 width=78 style='mso-width-source:userset;mso-width-alt:                                                               \n");
                htmlStr.Append("    					 2759;width:126pt'>                                                               \n");
                htmlStr.Append("    					 <col class=xl6617145 width=56 style='mso-width-source:userset;mso-width-alt:                                                               \n");
                htmlStr.Append("    					 1991;width:55pt'>                                                               \n");
                htmlStr.Append("    					 <col class=xl6617145 width=34 style='mso-width-source:userset;mso-width-alt:                                                               \n");
                htmlStr.Append("    					 1223;width:43pt'>                                                               \n");
                htmlStr.Append("    					 <col class=xl6617145 width=49 style='mso-width-source:userset;mso-width-alt:                                                               \n");
                htmlStr.Append("    					 1735;width:46pt'>                                                               \n");
                htmlStr.Append("    					 <col class=xl6617145 width=42 style='mso-width-source:userset;mso-width-alt:                                                               \n");
                htmlStr.Append("    					 1479;width:30pt'>                                                               \n");
                htmlStr.Append("    					 <col class=xl6617145 width=13 style='mso-width-source:userset;mso-width-alt:                                                               \n");
                htmlStr.Append("    					 455;width:16pt'>                                                               \n");
                htmlStr.Append("    					 <col class=xl6617145 width=49 style='mso-width-source:userset;mso-width-alt:                                                               \n");
                htmlStr.Append("    					 1735;width:82pt'>                                                               \n");
                htmlStr.Append("    					 <col class=xl6617145 width=78 style='mso-width-source:userset;mso-width-alt:                                                               \n");
                htmlStr.Append("    					 2759;width:96pt'>                                                               \n");
                htmlStr.Append("    					 <col class=xl6617145 width=13 style='mso-width-source:userset;mso-width-alt:                                                               \n");
                htmlStr.Append("    					 455;width:16pt'>                                                               \n");
                htmlStr.Append("    	 <tr height=33 style='mso-height-source:userset;height:25.05pt'>                                                               \n");
                htmlStr.Append("    	  <td height=33 class=xl14417145 style='height:25.05pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl6817145>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td align=left class=xl12917145 valign=top><span                                                               \n");
                htmlStr.Append("    					style='mso-ignore: vglayout; position: absolute; z-index: 1; margin-left: 5px; margin-top: 10px; width: 140px; height: 140px'><img                                                               \n");
                htmlStr.Append("    						width=140 height=140                                                               \n");
                htmlStr.Append("    						src='${ pageContext.request.contextPath}/assets/images/QR_code.png'                                                               \n");
                htmlStr.Append("    						alt='logovuong' v:shapes='Picture_x0020_2'></span>                                                               \n");
                htmlStr.Append("    					<![endif]><span style='mso-ignore: vglayout2'>                                                               \n");
                htmlStr.Append("    						<table cellpadding=0 cellspacing=0>                                                               \n");
                htmlStr.Append("    							<tr>                                                               \n");
                htmlStr.Append("    								<td height=46 width=29                                                               \n");
                htmlStr.Append("    									style='height: 35.1pt; width: 22pt'></td>                                                               \n");
                htmlStr.Append("    							</tr>                                                               \n");
                htmlStr.Append("    						</table>                                                               \n");
                htmlStr.Append("    				</span></td>                                                               \n");
                htmlStr.Append("    	  <td class=xl12917145>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td colspan=7 class=xl14817145>HÓA &#272;&#416;N BÁN HÀNG (<font                                                               \n");
                htmlStr.Append("    	  class='font2417145'>INVOICE</font><font class='font2317145'>)</font></td>                                                               \n");
                htmlStr.Append("    	  <td class=xl12917145>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl12917145>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl12917145>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl12917145>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl13017145>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	 </tr>                                                               \n");
                htmlStr.Append("    	 <tr height=21 style='height:15.6pt'>                                                               \n");
                htmlStr.Append("    	  <td height=21 class=xl7017145 style='height:15.6pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl6617145>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl6617145>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl6617145>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td colspan=7 class=xl8217145>(Dùng cho t&#7893; ch&#7913;c, cá nhân trong                                                               \n");
                htmlStr.Append("    	  khu phi thu&#7871; quan)</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl7117145 colspan=5 style='border-right:.5pt solid black'><font class='font1217145'></font><font class='font917145'></font><font                                                               \n");
                htmlStr.Append("    	  class='font817145'></font></td>                                                               \n");
                htmlStr.Append("    	 </tr>                                                               \n");
                htmlStr.Append("    	 <tr height=18 style='height:13.8pt'>                                                               \n");
                htmlStr.Append("    	  <td height=18 class=xl7017145 style='height:13.8pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl6617145>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl6617145>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl6617145>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td colspan=7 class=xl17817145><span style='mso-spacerun:yes'>(HÓA &#272;&#416;N CHUY&#7874;N &#272;&#7892;I T&#7914; HÓA &#272;&#416;N &#272;I&#7878;N T&#7916;)</span></td>                                                               \n");
                htmlStr.Append("    	  <td class=xl7117145 colspan=4>Ký hi&#7879;u (<font class='font1217145'>Serial</font><font                                                               \n");
                htmlStr.Append("    	  class='font917145'>):</font><font class='font817145'> " + dt.Rows[0]["templateCode"] + "" + dt.Rows[0]["InvoiceSerialNo"] + "</font></td>                                                               \n");
                htmlStr.Append("    	  <td class=xl7217145>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	 </tr>                                                               \n");
                htmlStr.Append("    	 <tr height=24 style='mso-height-source:userset;height:18.0pt'>                                                               \n");
                htmlStr.Append("    	  <td height=24 class=xl7017145 style='height:18.0pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl6617145>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl6617145>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl6617145>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td colspan=7 class=xl14917145 width=383 style='width:288pt'>Ngày <font                                                               \n");
                htmlStr.Append("    	  class='font1217145'>(Date)</font><font class='font1117145'> </font><font                                                               \n");
                htmlStr.Append("    	  class='font717145'>  " + dt.Rows[0]["invoiceissueddate_dd"] + " tháng </font><font class='font1217145'>(month)</font><font                                                               \n");
                htmlStr.Append("    	  class='font717145'> " + dt.Rows[0]["invoiceissueddate_mm"] + " n&#259;m </font><font class='font1217145'>(year)</font><font                                                               \n");
                htmlStr.Append("    	  class='font717145'> " + dt.Rows[0]["invoiceissueddate_yyyy"] + "</font></td>                                                               \n");
                htmlStr.Append("    	  <td class=xl7117145 colspan=4>S&#7889; (<font class='font1217145'>No</font><font                                                               \n");
                htmlStr.Append("    	  class='font917145'>.):</font><font class='font817145'><span                                                               \n");
                htmlStr.Append("    	  style='mso-spacerun:yes'>      </span></font><font class='font1417145'><span                                                               \n");
                htmlStr.Append("    	  style='mso-spacerun:yes'> </span></font><font class='font1617145'>" + dt.Rows[0]["InvoiceNumber"] + "</font></td>                                                               \n");
                htmlStr.Append("    	  <td class=xl7417145 width=13 style='width:10pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	 </tr>                                                               \n");
                htmlStr.Append("    	 <tr class=xl6817145 height=24 style='mso-height-source:userset;height:18.0pt'>                                                               \n");
                htmlStr.Append("    	  <td height=24 class=xl6717145 style='height:18.0pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl17917145 colspan=10>&#272;&#417;n v&#7883; bán hàng (<font                                                               \n");
                htmlStr.Append("    	  class='font1217145'>Company name</font><font class='font717145'>): " + dt.Rows[0]["Seller_Name"] + "</font><font                                                               \n");
                htmlStr.Append("    	  class='font517145'> </font></td>                                                               \n");
                htmlStr.Append("    	  <td class=xl6817145>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl6817145>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl6817145>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl6817145>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl6917145>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	 </tr>                                                               \n");
                htmlStr.Append("    	 <tr height=24 style='mso-height-source:userset;height:18.0pt'>                                                               \n");
                htmlStr.Append("    	  <td height=24 class=xl7017145 style='height:18.0pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl7517145 colspan=14>&#272;&#7883;a ch&#7881; (<font                                                               \n");
                htmlStr.Append("    	  class='font1217145'>Address</font><font class='font717145'>): </font><font                                                               \n");
                htmlStr.Append("    	  class='font2117145'>" + dt.Rows[0]["Seller_Address"] + "</font></td>                                                               \n");
                htmlStr.Append("    	  <td class=xl7217145></td>                                                               \n");
                htmlStr.Append("    	 </tr>                                                               \n");
                htmlStr.Append("    	 <tr height=24 style='mso-height-source:userset;height:18.0pt'>                                                               \n");
                htmlStr.Append("    	  <td height=24 class=xl7017145 style='height:18.0pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl7517145 colspan=15 style='border-right:.5pt solid black'>S&#7889;                                                               \n");
                htmlStr.Append("    	  tài kho&#7843;n (<font class='font1217145'>Acc. code</font><font                                                               \n");
                htmlStr.Append("    	  class='font717145'>): " + dt.Rows[0]["Seller_AccountNo"] + "/" + dt.Rows[0]["Seller_AccountNo2"] + " " + dt.Rows[0]["BANK_NM78"] + "</font></td>                                                               \n");
                htmlStr.Append("    	 </tr>                                                               \n");
                htmlStr.Append("    	 <tr height=24 style='mso-height-source:userset;height:18.0pt'>                                                               \n");
                htmlStr.Append("    	  <td height=24 class=xl7017145 style='height:18.0pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl7517145 colspan=6>&#272;i&#7879;n tho&#7841;i (<font                                                               \n");
                htmlStr.Append("    	  class='font1217145'>Tel</font><font class='font717145'>): " + dt.Rows[0]["Seller_Tel"] + " - Fax:                                                               \n");
                htmlStr.Append("    	  " + dt.Rows[0]["Seller_Fax"] + "</font></td>                                                               \n");
                htmlStr.Append("    	  <td class=xl7517145 colspan=7 >Mã                                                               \n");
                htmlStr.Append("    	  s&#7889; thu&#7871; (<font class='font1217145'>Tax code</font><font                                                               \n");
                htmlStr.Append("    	  class='font717145'>): </font><font class='font1517145'>" + dt.Rows[0]["Seller_TaxCode"] + "</font></td>                                                               \n");
                htmlStr.Append("    	  <td class=xl6617145 colspan=2 style='border-right:.5pt solid black'>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	 </tr>                                                               \n");



                htmlStr.Append("    	 <tr height=24 style='mso-height-source:userset;height:18.0pt'>                                                               \n");
                htmlStr.Append("    	  <td height=24 class=xl6717145 style='height:18.0pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl17917145 colspan=5>H&#7885; tên ng&#432;&#7901;i mua hàng (<font                                                               \n");
                htmlStr.Append("    	  class='font1217145'>Customer's name</font><font class='font717145'>):</font>" + dt.Rows[0]["buyer"] + "</td>                                                               \n");
                htmlStr.Append("    	  <td colspan=10 class=xl17617145 style='border-right:.5pt solid black'>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	 </tr>                                                               \n");
                htmlStr.Append("    	 <tr height=24 style='mso-height-source:userset;height:18.0pt'>                                                               \n");
                htmlStr.Append("    	  <td height=24 class=xl7017145 style='height:18.0pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl7517145 colspan=14>Tên &#273;&#417;n v&#7883; (<font                                                               \n");
                htmlStr.Append("    	  class='font1217145'>Company's name</font><font class='font717145'>): " + dt.Rows[0]["buyerlegalname"] + "</font></td>                                                               \n");

                htmlStr.Append("    	  <td class=xl14517145></td>                                                               \n");
                htmlStr.Append("    	 </tr>                                                               \n");
                htmlStr.Append("    	 <tr height=24 style='mso-height-source:userset;height:18.0pt'>                                                               \n");
                htmlStr.Append("    	  <td height=24 class=xl7017145 style='height:18.0pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl7517145 colspan=14>&#272;&#7883;a ch&#7881; (<font class='font1217145'>Address</font><font class='font717145'>): " + dt.Rows[0]["buyeraddress"] + "</font></td>                                                               \n");
                //htmlStr.Append("    	  <td class=xl7717145>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl7817145>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	 </tr>                                                               \n");
                htmlStr.Append("    	 <tr height=24 style='mso-height-source:userset;height:18.0pt'>                                                               \n");
                htmlStr.Append("    	  <td height=24 class=xl7017145 style='height:18.0pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl7517145 colspan=4>S&#7889; tài kho&#7843;n (<font                                                               \n");
                htmlStr.Append("    	  class='font1217145'>Account code</font><font class='font717145'>):" + dt.Rows[0]["BuyerAccountNo"] + "</font></td>                                                               \n");
                htmlStr.Append("    	  <td class=xl7917145>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl6617145>&nbsp;</td>                                                               \n");

                htmlStr.Append("    	  <td class=xl7517145>Mã s&#7889; thu&#7871;</font><font                                                               \n");
                htmlStr.Append("    	  class='font1017145'> (</font><font class='font1217145'>Tax code</font><font                                                               \n");
                htmlStr.Append("    	  class='font1017145'>): </font><font class='font1517145'>" + dt.Rows[0]["BuyerTaxCode"] + "</font></td>                                                               \n");
                htmlStr.Append("    	  <td class=xl6617145>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl6617145>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl6617145>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl6617145>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl6617145>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl6617145>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl6617145>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl7217145>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	 </tr>                                                               \n");

                htmlStr.Append("    	 <tr height=26 style='mso-height-source:userset;height:19.95pt'>                                                               \n");
                htmlStr.Append("    	  <td height=26 class=xl7017145 style='height:19.95pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl7517145 colspan=6>Hình th&#7913;c thanh toán (<font                                                               \n");
                htmlStr.Append("    	  class='font1217145'>Mod of payment</font><font class='font717145'>): " + dt.Rows[0]["PaymentMethodCK"] + "</font><font                                                               \n");
                htmlStr.Append("    	  class='font517145'></font></td>                                                               \n");
                //  htmlStr.Append("    	  <td class=xl7517145>&nbsp;</td>                                                               \n");
                //   htmlStr.Append("    	  <td class=xl6617145>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl6617145 colspan=7><font class='font717145'><font class='font717145'>Đơn vị tiền tệ<font class='font1217145'> </font><font class='font1217145'>(Currency) : " + dt.Rows[0]["CurrencyCodeUSD"] + "</font></font></td>                                                               \n");
                htmlStr.Append("    	  <td class=xl7217145 colspan=2>&nbsp;</td>                                                               \n");
                htmlStr.Append("    	 </tr>                                                               \n");
                htmlStr.Append("     <tr height=26 style='mso-height-source:userset;height:19.95pt'>                                                               \n");
                htmlStr.Append("      <td height=26 class=xl7017145 style='height:19.95pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("      <td class=xl7517145 colspan=6>Số hóa đơn thương mại (<font                                                               \n");
                htmlStr.Append("      class='font1217145'>Comercial Invoice No.</font><font class='font1217145'>):<span                                                               \n");
                htmlStr.Append("      style='mso-spacerun:yes'>" + dt.Rows[0]["attribute_03"] + "</span></font></td>                                                               \n");
                htmlStr.Append("      <td class=xl7517145>&nbsp;</td>                                                               \n");
                htmlStr.Append("      <td class=xl6617145 colspan=6></td>                                                               \n");
                htmlStr.Append("      <td class=xl6617145>&nbsp;</td>                                                               \n");
                htmlStr.Append("                                                                     \n");
                htmlStr.Append("      <td class=xl7217145>&nbsp;</td>                                                               \n");
                htmlStr.Append("     </tr>                                                               \n");
                /* htmlStr.Append("     <tr height=24 style='mso-height-source:userset;height:18.0pt'>                                                               \n");
                 htmlStr.Append("      <td height=24 class=xl7017145 style='height:18.0pt'>&nbsp;</td>                                                               \n");
                 htmlStr.Append("      <td class=xl7517145 colspan=4>S&#7889; tài kho&#7843;n (<font                                                               \n");
                 htmlStr.Append("      class='font1217145'>Account code</font><font class='font717145'></font><font                                                               \n");
                 htmlStr.Append("      class='font1017145'>): </font><font class='font1517145'></font></td>                                                               \n");
                 htmlStr.Append("      <td class=xl7917145>&nbsp;</td>                                                               \n");
                 htmlStr.Append("      <td class=xl6617145>&nbsp;</td>                                                               \n");
                 htmlStr.Append("      <td class=xl6617145>&nbsp;</td>                                                               \n");
                 htmlStr.Append("      <td class=xl7517145>Mã s&#7889; thu&#7871;</font><font                                                               \n");
                 htmlStr.Append("      class='font1017145'> (</font><font class='font1217145'>Tax code</font><font                                                               \n");
                 htmlStr.Append("      class='font1017145'> )</font>: <%=v_tax_code %></td>                                                               \n");
                 htmlStr.Append("      <td class=xl6617145>&nbsp;</td>                                                               \n");
                 htmlStr.Append("      <td class=xl6617145>&nbsp;</td>                                                               \n");
                 htmlStr.Append("      <td class=xl6617145>&nbsp;</td>                                                               \n");
                 htmlStr.Append("      <td class=xl6617145>&nbsp;</td>                                                               \n");
                 htmlStr.Append("      <td class=xl6617145>&nbsp;</td>                                                               \n");
                 htmlStr.Append("      <td class=xl6617145>&nbsp;</td>                                                               \n");
                 htmlStr.Append("      <td class=xl7217145>&nbsp;</td>                                                               \n");
                 htmlStr.Append("     </tr>                                                               \n");
                 htmlStr.Append("     <tr height=26 style='mso-height-source:userset;height:19.95pt'>                                                               \n");
                 htmlStr.Append("      <td height=26 class=xl7017145 style='height:19.95pt'>&nbsp;</td>                                                               \n");
                 htmlStr.Append("      <td class=xl7517145 colspan=6>Hình th&#7913;c thanh toán (<font                                                               \n");
                 htmlStr.Append("      class='font1217145'>Mod of payment</font><font class='font717145'>): </font><font                                                               \n");
                 htmlStr.Append("      class='font517145'></font><%=v_payment_method %></td>                                                               \n");
                 htmlStr.Append("      <td class=xl7517145>&nbsp;</td>                                                               \n");
                 htmlStr.Append("      <td class=xl6617145 colspan=6><font class='font717145'>Đơn vị tiền tệ<font class='font1217145'> </font><font class='font1217145'>(Currency) :</font></font></td>                                                               \n");
                 htmlStr.Append("      <td class=xl6617145>&nbsp;</td>                                                               \n");
                 htmlStr.Append("                                                                     \n");
                 htmlStr.Append("      <td class=xl7217145>&nbsp;</td>                                                               \n");
                 htmlStr.Append("     </tr>                                                               \n");*/



                htmlStr.Append("    	 <tr class=xl8217145 height=21 style='height:15.6pt'>                                                               \n");
                htmlStr.Append("    	  <td colspan=2 height=21 class=xl15017145 style='height:15.6pt'>STT</td>                                                               \n");
                htmlStr.Append("    	  <td colspan=5 class=xl15017145 style='border-right:.5pt solid black'>Tên hàng                                                               \n");
                htmlStr.Append("    	  hóa, d&#7883;ch v&#7909;</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl8117145>&#272;&#417;n v&#7883; tính</td>                                                               \n");
                htmlStr.Append("    	  <td colspan=2 class=xl15017145 style='border-right:.5pt solid black'>S&#7889;                                                               \n");
                htmlStr.Append("    	  l&#432;&#7907;ng</td>                                                               \n");
                htmlStr.Append("    	  <td colspan=3 class=xl8117145>&#272;&#417;n giá</td>                                                               \n");
                htmlStr.Append("    	  <td colspan=3 class=xl15017145 style='border-right:.5pt solid black'>Thành                                                               \n");
                htmlStr.Append("    	  ti&#7873;n</td>                                                               \n");
                htmlStr.Append("    	 </tr>                                                               \n");
                htmlStr.Append("    	 <tr class=xl8317145 height=18 style='height:13.2pt'>                                                               \n");
                htmlStr.Append("    	  <td colspan=2 height=18 class=xl16317145 style='height:13.2pt'>No.</td>                                                               \n");
                htmlStr.Append("    	  <td colspan=5 class=xl15217145 style='border-right:.5pt solid black'>Description                                                               \n");
                htmlStr.Append("    	  of goods</td>                                                               \n");
                htmlStr.Append("    	  <td class=xl8317145>Unit</td>                                                               \n");
                htmlStr.Append("    	  <td colspan=2 class=xl16317145 style='border-right:.5pt solid black'>Quatity</td>                                                               \n");
                htmlStr.Append("    	  <td colspan=3 class=xl8317145>Unit price</td>                                                               \n");
                htmlStr.Append("    	  <td colspan=3 class=xl16317145 style='border-right:.5pt solid black'>Amount</td>                                                               \n");
                htmlStr.Append("    	 </tr>                                                               \n");
                htmlStr.Append("    	 <tr class=xl8517145 height=20 style='mso-height-source:userset;height:15.0pt'>                                                               \n");
                htmlStr.Append("    	  <td colspan=2 height=20 class=xl15817145 style='height:15.0pt'>1</td>                                                               \n");
                htmlStr.Append("    	  <td colspan=5 height=20 width=291 style='border-right:.5pt solid black;                                                               \n");
                htmlStr.Append("    	  height:15.0pt;width:218pt' align=left valign=top><span style='mso-ignore:vglayout;                                                               \n");
                htmlStr.Append("    	  position:absolute;z-index:3;margin-left:106px;margin-top:12px;width:1100px;                                                               \n");
                htmlStr.Append("    	  height:900px'><img width=1100 height=900                                                               \n");
                htmlStr.Append("    	  src='${pageContext.request.contextPath}/assets/images/YUPOONG_002.png'                                                              \n");
                htmlStr.Append("    	  v:shapes='Picture_x0020_3'></span><![endif]><span style='mso-ignore:vglayout2'>                                                               \n");
                htmlStr.Append("    	  <table cellpadding=0 cellspacing=0>                                                               \n");
                htmlStr.Append("    	   <tr>                                                               \n");
                htmlStr.Append("    	    <td colspan=5 height=20 class=xl15817145 width=291 style='border-right:                                                               \n");
                htmlStr.Append("    	    .5pt solid black;border-right:none;border-top:none;border-bottom:none;height:15.0pt;width:218pt'>2</td>                                                               \n");
                htmlStr.Append("    	   </tr>                                                               \n");
                htmlStr.Append("    	  </table>                                                               \n");
                htmlStr.Append("    	  </span></td>                                                               \n");
                htmlStr.Append("    	  <td class=xl8417145>3</td>                                                               \n");
                htmlStr.Append("    	  <td colspan=2 class=xl15817145 style='border-right:.5pt solid black'>4</td>                                                               \n");
                htmlStr.Append("    	  <td colspan=3 class=xl8417145>5</td>                                                               \n");
                htmlStr.Append("    	  <td colspan=3 class=xl15817145 style='border-right:.5pt solid black'>4 x 5 =6</td>                                                               \n");
                htmlStr.Append("    	 </tr>                                                               \n");

                //////////////////////////////

                v_rowHeight = "48.0pt"; //"26.5pt";
                v_rowHeightEmpty = "22.0pt";
                v_rowHeightNumber = 28;

                v_rowHeightLast = "40.0pt";// "23.5pt";
                v_rowHeightLastNumber = 28;// 23.5;
                v_rowHeightEmptyLast = "23.5pt"; //"23.5pt";


                for (int dtR = 0; dtR < page[k]; dtR++)
                {
                    if (!vlongItemName && dt_d.Rows[v_index]["ITEM_NAME"].ToString().Length >= 92)
                    {
                        v_rowHeight = "55.0pt"; //"26.5pt";    
                        v_rowHeightLast = "40.0pt"; //"27.5pt";
                        v_rowHeightLastNumber = 28;//27.5;
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
                        htmlStr.Append("    <tr class=xl9317145 height=22 style='mso-height-source:userset;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                               \n");
                        htmlStr.Append("    									  <td colspan=2 height=22 class=xl9217145 style='border-right:.5pt solid black;                                                               \n");
                        htmlStr.Append("    									  height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ";text-align:center'>" + dt_d.Rows[v_index][7] + "</td>                                                               \n");
                        htmlStr.Append("    									  <td colspan=5 class=xl16017145 width=291 style='border-right:.5pt solid black;                                                               \n");
                        htmlStr.Append("    									  border-left:none;width:218pt'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                               \n");
                        htmlStr.Append("    									  <td class=xl8617145 style='border-left:none'>" + dt_d.Rows[v_index][1] + "</td>                                                               \n");
                        htmlStr.Append("    									  <td class=xl8717145 style='border-left:none;border-right:.5pt solid black' colspan=2>" + dt_d.Rows[v_index][2] + "&nbsp;&nbsp;</td>                                                               \n");
                        htmlStr.Append("    									  <td class=xl8917145 style='border-left:none;border-right:.5pt solid black' colspan=3>" + dt_d.Rows[v_index][3] + "&nbsp;&nbsp;&nbsp;&nbsp;</td>                                                               \n");
                        htmlStr.Append("    									  <td class=xl9217145 style='border-left:none' colspan=2>&nbsp;" + dt_d.Rows[v_index][4] + "</td>                                                               \n");
                        htmlStr.Append("    									  <td class=xl9117145><span style='mso-spacerun:yes'>   </span></td>                                                               \n");
                        htmlStr.Append("    									 </tr>                                                               \n");
                    }
                    else if (dtR == page[k] - 1)//dong cuoi moi trang
                    {
                        if (k < v_countNumberOfPages - 1) //trang giua
                        {
                            htmlStr.Append("    	<tr class=xl7517145 height=24 style='mso-height-source:userset;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                               \n");
                            htmlStr.Append("    								  <td colspan=2 height=24 class=xl12317145 width=39 style='border-right:.5pt solid black;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ";width:29pt'>" + dt_d.Rows[v_index][7] + "</td>                                                               \n");
                            htmlStr.Append("    								  <td colspan=5 class=xl12517145 width=166 style='width:125pt'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                               \n");
                            htmlStr.Append("    								  <td class=xl10117145 width=78 style='width:58pt'>" + dt_d.Rows[v_index][1] + "</td>                                                               \n");
                            htmlStr.Append("    								  <td class=xl10217145 width=56 colspan=2 style='border-left:none;width:42pt'>" + dt_d.Rows[v_index][2] + "&nbsp;&nbsp;</td>                                                               \n");

                            htmlStr.Append("    								  <td colspan=2 class=xl9517145 style='border-left:none'>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                                               \n");
                            htmlStr.Append("    								  <td class=xl10417145 width=13 style='width:10pt'></td>                                                               \n");
                            htmlStr.Append("    								  <td colspan=2 class=xl9517145 style='border-left:none'>" + dt_d.Rows[v_index][4] + "</td>                                                               \n");
                            htmlStr.Append("    								  <td class=xl10517145 width=13 style='width:10pt'></td>                                                               \n");
                            htmlStr.Append("    								 </tr>	                                                               \n");
                        }
                        else // trang cuoi
                        {
                            if (dtR == rowsPerPage - 1) // du 11 dong
                            {
                                htmlStr.Append("    	<tr class=xl7517145 height=24 style='mso-height-source:userset;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                               \n");
                                htmlStr.Append("    								  <td colspan=2 height=24 class=xl12317145 width=39 style='border-right:.5pt solid black;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ";width:29pt'>" + dt_d.Rows[v_index][7] + "</td>                                                               \n");
                                htmlStr.Append("    								  <td colspan=5 class=xl12517145 width=166 style='width:125pt'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                               \n");
                                htmlStr.Append("    								  <td class=xl10117145 width=78 style='width:58pt'>" + dt_d.Rows[v_index][1] + "</td>                                                               \n");
                                htmlStr.Append("    								  <td class=xl10217145 width=56 colspan=2 style='border-left:none;width:42pt'>" + dt_d.Rows[v_index][2] + "&nbsp;&nbsp;</td>                                                               \n");

                                htmlStr.Append("    								  <td colspan=2 class=xl9517145 style='border-left:none'>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                                               \n");
                                htmlStr.Append("    								  <td class=xl10417145 width=13 style='width:10pt'></td>                                                               \n");
                                htmlStr.Append("    								  <td colspan=2 class=xl9517145 style='border-left:none'>" + dt_d.Rows[v_index][4] + "</td>                                                               \n");
                                htmlStr.Append("    								  <td class=xl10517145 width=13 style='width:10pt'></td>                                                               \n");
                                htmlStr.Append("    								 </tr>	                                                               \n");
                            }
                            else
                            {
                                htmlStr.Append("    	<tr class=xl7517145 height=24 style='mso-height-source:userset;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                               \n");
                                htmlStr.Append("    								  <td colspan=2 height=24 class=xl12317145 width=39 style='border-right:.5pt solid black;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ";width:29pt'>" + dt_d.Rows[v_index][7] + "</td>                                                               \n");
                                htmlStr.Append("    								  <td colspan=5 class=xl12517145 width=166 style='width:125pt'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                               \n");
                                htmlStr.Append("    								  <td class=xl10117145 width=78 style='width:58pt'>" + dt_d.Rows[v_index][1] + "</td>                                                               \n");
                                htmlStr.Append("    								  <td class=xl10217145 width=56 colspan=2 style='border-left:none;width:42pt'>" + dt_d.Rows[v_index][2] + "&nbsp;&nbsp;</td>                                                               \n");

                                htmlStr.Append("    								  <td colspan=2 class=xl9517145 style='border-left:none'>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                                               \n");
                                htmlStr.Append("    								  <td class=xl10417145 width=13 style='width:10pt'></td>                                                               \n");
                                htmlStr.Append("    								  <td colspan=2 class=xl9517145 style='border-left:none'>" + dt_d.Rows[v_index][4] + "</td>                                                               \n");
                                htmlStr.Append("    								  <td class=xl10517145 width=13 style='width:10pt'></td>                                                               \n");
                                htmlStr.Append("    								 </tr>	                                                               \n");
                            }

                        }
                    }
                    else
                    { // dong giua                                                                                                                                    
                        htmlStr.Append("     <tr class=xl8217145 height=24 style='mso-height-source:userset;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                               \n");
                        htmlStr.Append("    								  <td colspan=2 height=24 class=xl12317145 width=39 style='border-right:.5pt solid black;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ";width:29pt'>" + dt_d.Rows[v_index][7] + "</td>                                                               \n");
                        htmlStr.Append("    								  <td colspan=5 class=xl13117145 width=291 style='border-right:.5pt solid black;                                                               \n");
                        htmlStr.Append("    								  border-left:none;width:218pt'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                               \n");
                        htmlStr.Append("    								  <td class=xl9417145 style='border-left:none'>" + dt_d.Rows[v_index][1] + "</td>                                                               \n");
                        htmlStr.Append("    								  <td class=xl9617145 style='border-left:none'colspan=2>" + dt_d.Rows[v_index][2] + "&nbsp;&nbsp;</td>                                                               \n");
                        htmlStr.Append("    								  <td colspan=2 class=xl9517145 style='border-left:none' >" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                                               \n");
                        htmlStr.Append("    								  <td class=xl9717145>&nbsp;</td>                                                               \n");
                        htmlStr.Append("    								  <td colspan=2 class=xl9517145 style='border-left:none'>" + dt_d.Rows[v_index][4] + "</td>                                                               \n");
                        htmlStr.Append("    								  <td class=xl9717145>&nbsp;</td>                                                               \n");
                        htmlStr.Append("    								 </tr>                                                               \n");
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
                    v_spacePerPage = 80;
                }

                if (k == v_countNumberOfPages - 1 && page[k] < rowsPerPage) // Trang cuoi khong du dong
                {
                    v_rowHeightEmptyLast = Math.Round(v_totalHeightLastPage / (rowsPerPage - page[k]), 2).ToString() + "pt";
                    for (int i = 0; i < rowsPerPage - page[k]; i++)
                    {
                        if (i == (rowsPerPage - page[k] - 1))
                        {
                            htmlStr.Append("    	<tr class=xl8217145 height=24 style='mso-height-source:userset;height:" + v_rowHeightEmptyLast + "'>                                                               \n");
                            htmlStr.Append("    												  <td colspan=2 height=24 class=xl12317145 width=39 style='border-right:.5pt solid black;                                                               \n");
                            htmlStr.Append("    												  height:" + v_rowHeightEmptyLast + ";width:29pt'></td>                                                               \n");
                            htmlStr.Append("    												  <td colspan=5 class=xl13117145 width=291 style='border-right:.5pt solid black;                                                               \n");
                            htmlStr.Append("    												  border-left:none;width:218pt'></td>                                                               \n");
                            htmlStr.Append("    												  <td class=xl9417145 style='border-left:none'></td>                                                               \n");
                            htmlStr.Append("    												  <td class=xl9517145 style='border-left:none'></td>                                                               \n");
                            htmlStr.Append("    												  <td class=xl9617145>&nbsp;</td>                                                               \n");
                            htmlStr.Append("    												  <td colspan=2 class=xl9517145 style='border-left:none'></td>                                                               \n");
                            htmlStr.Append("    												  <td class=xl9717145>&nbsp;</td>                                                               \n");
                            htmlStr.Append("    												  <td colspan=2 class=xl9517145 style='border-left:none'></td>                                                               \n");
                            htmlStr.Append("    												  <td class=xl9717145>&nbsp;</td>                                                               \n");
                            htmlStr.Append("    												 </tr>                                                               \n");
                        }
                        else
                        {
                            htmlStr.Append("    	<tr class=xl8217145 height=24 style='mso-height-source:userset;height:" + v_rowHeightEmptyLast + "'>                                                               \n");
                            htmlStr.Append("    												  <td colspan=2 height=24 class=xl12317145 width=39 style='border-right:.5pt solid black;                                                               \n");
                            htmlStr.Append("    												  height:" + v_rowHeightEmptyLast + ";width:29pt'></td>                                                               \n");
                            htmlStr.Append("    												  <td colspan=5 class=xl13117145 width=291 style='border-right:.5pt solid black;                                                               \n");
                            htmlStr.Append("    												  border-left:none;width:218pt'></td>                                                               \n");
                            htmlStr.Append("    												  <td class=xl9417145 style='border-left:none'></td>                                                               \n");
                            htmlStr.Append("    												  <td class=xl9517145 style='border-left:none'></td>                                                               \n");
                            htmlStr.Append("    												  <td class=xl9617145>&nbsp;</td>                                                               \n");
                            htmlStr.Append("    												  <td colspan=2 class=xl9517145 style='border-left:none'></td>                                                               \n");
                            htmlStr.Append("    												  <td class=xl9717145>&nbsp;</td>                                                               \n");
                            htmlStr.Append("    												  <td colspan=2 class=xl9517145 style='border-left:none'></td>                                                               \n");
                            htmlStr.Append("    												  <td class=xl9717145>&nbsp;</td>                                                               \n");
                            htmlStr.Append("    												 </tr>                                                               \n");
                        }
                    } // for

                }//Trang cuoi 11 dong

                if (k < v_countNumberOfPages - 1)
                {
                    htmlStr.Append("    	<tr class=xl8217145 height=24 style='mso-height-source:userset;height:" + (v_spacePerPage).ToString() + "pt;border-bottom:.5pt solid black;border-top:.5pt solid black;'>                                                               \n");
                    htmlStr.Append("    												  <td colspan=2 height=24 class=xl12317145 width=39 style='border-right:none;border-bottom:none;                                                               \n");
                    htmlStr.Append("    												  height:" + (v_spacePerPage).ToString() + "pt;width:29pt'></td>                                                               \n");
                    htmlStr.Append("    												  <td colspan=5 class=xl13117145 width=291 style='border-right:none;border-bottom:none;                                                               \n");
                    htmlStr.Append("    												  border-left:none;width:218pt'></td>                                                               \n");
                    htmlStr.Append("    												  <td class=xl9417145 style='border-right:none;border-left:none;border-bottom:none;'></td>                                                               \n");
                    htmlStr.Append("    												  <td class=xl9517145 style='border-right:none;border-left:none;border-bottom:none;'></td>                                                               \n");
                    htmlStr.Append("    												  <td class=xl9617145 style='border-right:none;border-bottom:none;'>&nbsp;</td>                                                               \n");
                    htmlStr.Append("    												  <td colspan=2 class=xl9517145 style='border-left:none;border-bottom:none;'></td>                                                               \n");
                    htmlStr.Append("    												  <td class=xl9717145 style='border-right:none;border-bottom:none;'>&nbsp;</td>                                                               \n");
                    htmlStr.Append("    												  <td colspan=2 class=xl9517145 style='border-left:none;border-bottom:none;'></td>                                                               \n");
                    htmlStr.Append("    												  <td class=xl9717145>&nbsp;</td>                                                               \n");
                    htmlStr.Append("    												 </tr>                                                               \n");
                    htmlStr.Append("	<table  border=0>                                                                                                                                                                                                 \n");
                    htmlStr.Append("		<tr height=20  style='height: 90pt'>                                                                                                                                                                \n");

                    htmlStr.Append("		</tr>      																																														\n");
                    htmlStr.Append("	</table>             																																										\n");
                }


            }// for k                                                                                                                             

            /////////////////
            htmlStr.Append("    	 <tr class=xl7517145 height=29 style='mso-height-source:userset;height:22.05pt'>                                                               \n");
            htmlStr.Append("    					  <td colspan=4  height=29 class=xl12717145 width=6 style='height:22.05pt;width:4pt'>&nbsp;Tỷ giá: " + dt.Rows[0]["exchangerate"] + "</td>                                                               \n");
            htmlStr.Append("    					  <td colspan=8 class=xl16917145 width=583 style='width:437pt'>C&#7897;ng                                                               \n");
            htmlStr.Append("    					  ti&#7873;n bán hàng hóa, d&#7883;ch v&#7909; (<font class='font1217145'>Total                                                               \n");
            htmlStr.Append("    					  amount</font><font class='font917145'>):</font></td>                                                               \n");
            htmlStr.Append("    					  <td class=xl12817145 width=13 style='width:10pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td colspan=2 class=xl16717145 style='border-left:none'>" + dt.Rows[0]["netamount_display"] + "</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl10717145>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					 </tr>                                                               \n");
            if (read_prive.ToString().Length >= 70)
            {
                htmlStr.Append("    					 <tr class=xl7517145 height=24 style='mso-height-source:userset;height:18.0pt'>                                                               \n");
                htmlStr.Append("    					  <td height=24 class=xl10617145 style='height:18.0pt;border-bottom:.5pt solid windowtext;'>&nbsp;</td>                                                               \n");
                htmlStr.Append("    					  <td colspan=14 class=xl17517145>S&#7889; ti&#7873;n vi&#7871;t b&#7857;ng                                                               \n");
                htmlStr.Append("    					  ch&#7919; (<font class='font1217145'>In words</font><font class='font717145'>):                                                               \n");
                htmlStr.Append("    					  </font><font class='font1317145'>" + read_prive + "</font></td>                                                               \n");
                htmlStr.Append("    					  <td class=xl10817145 width=13 style='width:10pt;border-bottom:.5pt solid windowtext'>&nbsp;</td>                                                               \n");
                htmlStr.Append("    					 </tr>                                                               \n");
            }
            else
            {

                htmlStr.Append("    					  <tr class=xl7517145 height=24 style='mso-height-source:userset;height:18.0pt'>                                                               \n");
                htmlStr.Append("    					  <td height=24 class=xl10617145 style='height:18.0pt;border-top:none'>&nbsp;</td>                                                               \n");
                htmlStr.Append("    					  <td colspan=14 class=xl17517145>S&#7889; ti&#7873;n vi&#7871;t b&#7857;ng                                                               \n");
                htmlStr.Append("    					  ch&#7919; (<font class='font1217145'>In words</font><font class='font717145'>):                                                               \n");
                htmlStr.Append("    					  </font><font class='font1317145'>" + read_prive + "</font></td>                                                               \n");
                htmlStr.Append("    					  <td class=xl10817145 width=13 style='width:10pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("    					 </tr>                                                               \n");
                htmlStr.Append("    					 <tr height=18 style='mso-height-source:userset;height:13.95pt'>                                                               \n");
                htmlStr.Append("    					  <td height=18 class=xl8017145 style='height:13.95pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("    					  <td class=xl10917145>&nbsp;</td>                                                               \n");
                htmlStr.Append("    					  <td class=xl10917145>&nbsp;</td>                                                               \n");
                htmlStr.Append("    					  <td class=xl10917145>&nbsp;</td>                                                               \n");
                htmlStr.Append("    					  <td class=xl11017145 width=41 style='width:31pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("    					  <td class=xl11017145 width=98 style='width:74pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("    					  <td class=xl11017145 width=27 style='width:20pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("    					  <td class=xl11017145 width=78 style='width:58pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("    					  <td class=xl11017145 width=56 style='width:42pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("    					  <td class=xl11017145 width=34 style='width:26pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("    					  <td class=xl11017145 width=49 style='width:37pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("    					  <td class=xl11017145 width=42 style='width:31pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("    					  <td class=xl11017145 width=13 style='width:10pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("    					  <td class=xl11017145 width=49 style='width:37pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("    					  <td class=xl11017145 width=78 style='width:58pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("    					  <td class=xl11117145 width=13 style='width:10pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("    					 </tr>                                                               \n");
            }




            htmlStr.Append("    <tr class=xl7717145 height=24 style='mso-height-source:userset;height:18.0pt'>                                                               \n");
            htmlStr.Append("      <td height=24 class=xl11317145 style='height:18.0pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td colspan=4 class=xl17417145 width=199 style='width:149pt'>Converter</td>                                                               \n");
            htmlStr.Append("      <td colspan=3 class=xl17417145 width=203 style='width:152pt'>Ng&#432;&#7901;i                                                               \n");
            htmlStr.Append("      mua hàng (Buyer)</td>                                                               \n");
            htmlStr.Append("      <td colspan=7 class=xl17417145 width=321 style='width:241pt'>Ng&#432;&#7901;i                                                               \n");
            htmlStr.Append("      mua hàng (Seller)</td>                                                               \n");
            htmlStr.Append("      <td class=xl11217145 width=13 style='width:10pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("     </tr>                                                               \n");
            htmlStr.Append("     <tr class=xl7117145 height=18 style='mso-height-source:userset;height:13.95pt'>                                                               \n");
            htmlStr.Append("      <td height=18 class=xl11417145 style='height:13.95pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td colspan=4 class=xl11517145 width=199 style='width:149pt'>(Ký, ghi rõ                                                               \n");
            htmlStr.Append("      h&#7885; tên)</td>                                                               \n");
            htmlStr.Append("      <td colspan=3 class=xl11517145 width=203 style='width:152pt'>(Ký, ghi rõ                                                               \n");
            htmlStr.Append("      h&#7885; tên)</td>                                                               \n");
            htmlStr.Append("      <td colspan=7 class=xl11517145 width=321 style='width:241pt'>(Ký, &#273;óng                                                               \n");
            htmlStr.Append("      d&#7845;u, ghi rõ h&#7885; tên)</td>                                                               \n");
            htmlStr.Append("      <td class=xl11617145 width=13 style='width:10pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("     </tr>                                                               \n");
            htmlStr.Append("     <tr class=xl11917145 height=20 style='mso-height-source:userset;height:15.45pt'>                                                               \n");
            htmlStr.Append("      <td height=20 class=xl11717145 style='height:15.45pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td colspan=4 class=xl17317145 width=199 style='width:149pt'>(Signature &amp;                                                               \n");
            htmlStr.Append("      full name)</td>                                                               \n");
            htmlStr.Append("      <td colspan=3 class=xl17317145 width=203 style='width:152pt'>(Signature &amp;                                                               \n");
            htmlStr.Append("      full name)</td>                                                               \n");
            htmlStr.Append("      <td colspan=7 class=xl17317145 width=321 style='width:241pt'>(Signature,                                                               \n");
            htmlStr.Append("      stamp &amp; full name)</td>                                                               \n");
            htmlStr.Append("      <td class=xl11817145 width=13 style='width:10pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("     </tr>                                                               \n");
            htmlStr.Append("    					 <tr height=20 style='mso-height-source:userset;height:15.45pt'>                                                               \n");
            htmlStr.Append("    					  <td height=20 class=xl7017145 style='height:15.45pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl6617145>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl6617145>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl6617145>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl7317145 width=41 style='width:31pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl7317145 width=98 style='width:74pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl7317145 width=27 style='width:20pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl7317145 width=78 style='width:58pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl7317145 width=56 style='width:42pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl7317145 width=34 style='width:26pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl7317145 width=49 style='width:37pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl7317145 width=42 style='width:31pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl7317145 width=13 style='width:10pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl7317145 width=49 style='width:37pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl7317145 width=78 style='width:58pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl12017145 width=13 style='width:10pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					 </tr>                                                               \n");
            htmlStr.Append("    					 <tr height=22 style='mso-height-source:userset;height:16.5pt'>                                                               \n");
            htmlStr.Append("    					  <td height=22 class=xl7017145 style='height:16.5pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl6617145>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl6617145>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl6617145>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl7317145 width=41 style='width:31pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl7317145 width=98 style='width:74pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl7317145 width=27 style='width:20pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl7317145 width=78 style='width:58pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl13317145 colspan=3>Signature Valid</td>                                                               \n");
            htmlStr.Append("    					                                                                 \n");
            if (dt.Rows[0]["sign_yn"].ToString() == "Y")
            {

                htmlStr.Append("    					  <td align=left class=xl13417145 valign=top><span style='mso-ignore:vglayout;                                                               \n");
                htmlStr.Append("    					  position:absolute;z-index:2;margin-left:18px;margin-top:7px;width:82px;                                                               \n");
                htmlStr.Append("    					  height:54px'><img width=82 height=54                                                              \n");
                htmlStr.Append("    					  src='${ pageContext.request.contextPath}/assets/img/check_signed.png'                                                               \n");
                htmlStr.Append("    					  v:shapes='Picture_x0020_8'></span><![endif]><span style='mso-ignore:vglayout2'>                                                               \n");
                htmlStr.Append("    					  <table cellpadding=0 cellspacing=0>                                                               \n");
                htmlStr.Append("    					   <tr>                                                               \n");
                htmlStr.Append("    					    <td height=22 width=42 style='height:16.5pt;width:31pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("    					   </tr>                                                               \n");
                htmlStr.Append("    					  </table>                                                               \n");
                htmlStr.Append("    					  </span></td>                                                               \n");
            }
            else
            {
                htmlStr.Append("    								<td height=22 class=xl13417145 width=42 style='height:16.5pt;width:31pt'>&nbsp;</td>                                                               \n");
            }


            htmlStr.Append("    					  <td class=xl13517145>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl13517145>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl13617145>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl13817145>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					 </tr>                                                               \n");

            htmlStr.Append("    					 <tr height=20 style='mso-height-source:userset;height:15.45pt'>                                                               \n");
            htmlStr.Append("    					  <td height=20 class=xl7017145 style='height:15.45pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl6617145>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl6617145>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl6617145>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl7317145 width=41 style='width:31pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl7317145 width=98 style='width:74pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl7317145 width=27 style='width:20pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl7317145 width=78 style='width:58pt'>&nbsp;</td>                                                               \n");

            if (dt.Rows[0]["sign_yn"].ToString() == "Y")
            {
                htmlStr.Append("    					  <td colspan=7 class=xl17017145 style='border-right:.5pt solid black'><font                                                               \n");
                htmlStr.Append("    					  class='font1717145'>&#272;&#432;&#7907;c ký b&#7903;i:</font><font                                                               \n");
                htmlStr.Append("    					  class='font1817145'> </font><font class='font1917145'>" + dt.Rows[0]["SignedBy"] + "</font></td>                                                               \n");
            }
            else
            {
                htmlStr.Append("    								<td colspan=7 class=xl17017145 style='border-right:.5pt solid black'></td>                                                               \n");
            }

            htmlStr.Append("                         <td class=xl13917145 width= 13 style='width:10pt'>&nbsp;</td>    \n");
            htmlStr.Append("    					 </tr>                                                               \n");
            htmlStr.Append("    					 <tr height=21 style='height:15.6pt'>                                                               \n");
            htmlStr.Append("    					  <td height=21 class=xl7017145 style='height:15.6pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl6517145></td>                                                               \n");
            htmlStr.Append("    					  <td class=xl6617145>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl6617145>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl7317145 width=41 style='width:31pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl7317145 width=98 style='width:74pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl7317145 width=27 style='width:20pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl7317145 width=78 style='width:58pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl14017145 colspan=3>Ngày Ký: <font class='font2017145'>" + dt.Rows[0]["SignedDate"] + "</font></td>                                                               \n");
            htmlStr.Append("    					  <td class=xl14117145>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl14117145>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl14117145>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl14217145>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl13717145>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					 </tr>                                                               \n");
            htmlStr.Append("    	 <tr height=16 style='mso-height-source:userset;height:2.0pt'>                                                               \n");
            htmlStr.Append("      <td height=16 class=xl70171451 style='height:2.0pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td colspan=14 class=xl166171451></td>                                                               \n");
            htmlStr.Append("      <td class=xl120171451 width=13 style='width:10pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("     </tr>                                                               \n");
            htmlStr.Append("                                                                    \n");
         /*   htmlStr.Append("    	 <tr height=22 style='mso-height-source:userset;height:16.5pt'>                                                               \n");
            htmlStr.Append("      <td height=22 class=xl7017145 style='height:16.5pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl6617145>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl6617145>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl6617145>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl7317145 width=41 style='width:31pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl7317145 width=98 style='width:74pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl7317145 width=27 style='width:20pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl7317145 width=78 style='width:58pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl13317145 colspan=3>Signature Valid</td>                                                               \n");
            if (dt.Rows[0]["sign_yn"].ToString() == "Y")
            {

                htmlStr.Append("      <td align=left valign=top class=xl13417145><span style='mso-ignore:vglayout;                                                               \n");
                htmlStr.Append("      position:absolute;z-index:2;margin-left:18px;margin-top:7px;width:82px;                                                               \n");
                htmlStr.Append("      height:54px'><img width=82 height=54                                                               \n");
                htmlStr.Append("      src='${pageContext.request.contextPath}/assets/img/check_signed.png'                                                               \n");
                htmlStr.Append("      v:shapes='Picture_x0020_8'></span><![endif]><span style='mso-ignore:vglayout2'>                                                               \n");
                htmlStr.Append("      <table cellpadding=0 cellspacing=0>                                                               \n");
                htmlStr.Append("       <tr>                                                               \n");
                htmlStr.Append("        <td height=22  width=42 style='height:16.5pt;width:31pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("       </tr>                                                               \n");
                htmlStr.Append("      </table>                                                               \n");
                htmlStr.Append("      </span></td>                                                               \n");
            }
            else
            {
                htmlStr.Append("    								<td height=22 class=xl13417145 width=42 style='height:16.5pt;width:31pt'>&nbsp;</td>                                                               \n");
            }
            htmlStr.Append("      <td class=xl13517145>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl13517145>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl13617145>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl13817145>&nbsp;</td>                                                               \n");
            htmlStr.Append("     </tr>                                                               \n");
            htmlStr.Append("     <tr height=20 style='mso-height-source:userset;height:15.45pt'>                                                               \n");
            htmlStr.Append("      <td height=20 class=xl7017145 style='height:15.45pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl6617145>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl6617145>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl6617145>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl7317145 width=41 style='width:31pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl7317145 width=98 style='width:74pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl7317145 width=27 style='width:20pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl7317145 width=78 style='width:58pt'>&nbsp;</td>                                                               \n");
            if (dt.Rows[0]["sign_yn"].ToString() == "Y")
            {
                htmlStr.Append("      <td colspan=7 class=xl17017145 style='border-right:.5pt solid black'><font                                                               \n");
                htmlStr.Append("      class='font1717145'>&#272;&#432;&#7907;c ký b&#7903;i:</font><font                                                               \n");
                htmlStr.Append("      class='font1817145'> </font><font class='font1917145'>TỔNG CỤC THUẾ</font></td>                                                               \n");
            }
            else
            {
                htmlStr.Append("    								<td colspan=7 class=xl17017145 style='border-right:.5pt solid black'></td>                                                               \n");
            }
            htmlStr.Append("      <td class=xl13917145 width=13 style='width:10pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("     </tr>                                                               \n");
            htmlStr.Append("     <tr height=21 style='height:15.6pt'>                                                               \n");
            htmlStr.Append("      <td height=21 class=xl7017145 style='height:15.6pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl6517145></td>                                                               \n");
            htmlStr.Append("      <td class=xl6617145>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl6617145>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl7317145 width=41 style='width:31pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl7317145 width=98 style='width:74pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl7317145 width=27 style='width:20pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl7317145 width=78 style='width:58pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl14017145 colspan=3>Ngày Ký: <font class='font2017145'>" + dt.Rows[0]["SignedDate"] + "</font></td>                                                               \n");
            htmlStr.Append("      <td class=xl14117145>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl14117145>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl14117145>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl14217145>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl13717145>&nbsp;</td>                                                               \n");
            htmlStr.Append("     </tr>                                                               \n");
            */
            htmlStr.Append("    					 <tr height=18 style='mso-height-source:userset;height:13.95pt'>                                                               \n");
            htmlStr.Append("    	  <td height=18 class=xl7017145 style='height:13.95pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    	  <td class=xl14317145 colspan=3>Mã CQT: " + dt.Rows[0]["cqt_mccqt_id"] + "</td>                                                               \n");
            htmlStr.Append("    	  <td class=xl7317145 width=41 style='width:31pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    	  <td class=xl7317145 width=98 style='width:74pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    	  <td class=xl7317145 width=27 style='width:20pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    	  <td class=xl7317145 width=78 style='width:58pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    	  <td class=xl7317145 width=56 style='width:42pt'>Mã nh&#7853;n hóa &#273;&#417;n: " + dt.Rows[0]["matracuu"] + "&nbsp;</td>                                                               \n");
            htmlStr.Append("    	  <td class=xl7317145 width=34 style='width:26pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    	  <td class=xl7317145 width=49 style='width:37pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    	  <td class=xl7317145 width=42 style='width:31pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    	  <td class=xl7317145 width=13 style='width:10pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    	  <td class=xl7317145 width=49 style='width:37pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    	  <td class=xl7317145 width=78 style='width:58pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    	  <td class=xl12017145 width=13 style='width:10pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    	 </tr>                                                               \n");
            htmlStr.Append("    					 <tr height=18 style='mso-height-source:userset;height:13.95pt'>                                                               \n");
            htmlStr.Append("    					  <td height=18 class=xl7017145 style='height:13.95pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl6617145 colspan=7>Tra c&#7913;u t&#7841;i Website: <font                                                               \n");
            htmlStr.Append("    					  class='font617145'><span style='mso-spacerun:yes'> </span></font><font                                                               \n");
            htmlStr.Append("    					  class='font2217145'>" + dt.Rows[0]["WEBSITE_EI"] + "</font></td>                                                               \n");
            htmlStr.Append("    					  <td class=xl7317145 width=56 style='width:42pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl7317145 width=34 style='width:26pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl7317145 width=49 style='width:37pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl7317145 width=42 style='width:31pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl7317145 width=13 style='width:10pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl7317145 width=49 style='width:37pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl7317145 width=78 style='width:58pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl12017145 width=13 style='width:10pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					 </tr>                                                               \n");
            htmlStr.Append("    					 <tr height=16 style='mso-height-source:userset;height:12.0pt'>                                                               \n");
            htmlStr.Append("    					  <td height=16 class=xl7017145 style='height:12.0pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td colspan=14 class=xl16617145>(C&#7847;n ki&#7875;m tra, &#273;&#7889;i                                                               \n");
            htmlStr.Append("    					  chi&#7871;u khi l&#7853;p, giao nh&#7853;n hóa &#273;&#417;n)</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl12017145 width=13 style='width:10pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					 </tr>                                                               \n");
            htmlStr.Append("    					 <tr height=16 style='mso-height-source:userset;height:12.0pt'>                                                               \n");
            htmlStr.Append("    					  <td height=16 class=xl8017145 style='height:12.0pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					  <td colspan=14 class=xl15317145>" + dt.Rows[0]["CONTRACT_INFO_EI"] + "</td>                                                               \n");
            htmlStr.Append("    					  <td class=xl12217145 width=13 style='width:10pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    					 </tr>                                                               \n");
            htmlStr.Append("    					 <![if supportMisalignedColumns]>                                                               \n");
            htmlStr.Append("    					 <tr height=0 style='display:none'>                                                               \n");
            htmlStr.Append("    					  <td width=6 style='width:4pt'></td>                                                               \n");
            htmlStr.Append("    					  <td width=33 style='width:25pt'></td>                                                               \n");
            htmlStr.Append("    					  <td width=70 style='width:52pt'></td>                                                               \n");
            htmlStr.Append("    					  <td width=55 style='width:41pt'></td>                                                               \n");
            htmlStr.Append("    					  <td width=41 style='width:31pt'></td>                                                               \n");
            htmlStr.Append("    					  <td width=98 style='width:74pt'></td>                                                               \n");
            htmlStr.Append("    					  <td width=27 style='width:20pt'></td>                                                               \n");
            htmlStr.Append("    					  <td width=78 style='width:58pt'></td>                                                               \n");
            htmlStr.Append("    					  <td width=56 style='width:42pt'></td>                                                               \n");
            htmlStr.Append("    					  <td width=34 style='width:26pt'></td>                                                               \n");
            htmlStr.Append("    					  <td width=49 style='width:37pt'></td>                                                               \n");
            htmlStr.Append("    					  <td width=42 style='width:31pt'></td>                                                               \n");
            htmlStr.Append("    					  <td width=13 style='width:10pt'></td>                                                               \n");
            htmlStr.Append("    					  <td width=49 style='width:37pt'></td>                                                               \n");
            htmlStr.Append("    					  <td width=78 style='width:58pt'></td>                                                               \n");
            htmlStr.Append("    					  <td width=13 style='width:10pt'></td>                                                               \n");
            htmlStr.Append("    					 </tr>                                                               \n");
            htmlStr.Append("    					 <![endif]>                                                               \n");
            htmlStr.Append("    					</table>                                                               \n");

            htmlStr.Append("    					</body>                                                               \n");
            htmlStr.Append("    </html>                                                               \n");
            ////////////////

            string filePath = "";

            using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"D:\GIT\EInvoice\PDF\" + tei_einvoice_m_pk + ".html"))
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